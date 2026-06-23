"""
Moteur de calcul TAFIROHA.

Reproduit les formules des feuilles BILAN, RESULTAT et TFT du classeur
TAFIROHA-DGI-FR (systeme normal OHADA) a partir d'une balance comptable a 6
colonnes (Compte, BE_Debit, BE_Credit, Mvt_Debit, Mvt_Credit, BS_Debit,
BS_Credit), comme le fait la feuille "TABLE DE CONVERSION" du fichier Excel
original (tables BalanceN / BalanceN1), et d'un detail de comptes a eclater
(106, 154, 1984, 4781, 4791, 4816, 4817, 4818) comme le fait la feuille
"Table TFT" (tables TFTN / TFTN1), necessaire au TFT.
"""
import json
import re
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, "sheets_raw.json"), encoding="utf-8") as fh:
    SHEETS_RAW = json.load(fh)

# Couvre SUMIFS(BalanceN[...], BalanceN[Table], "op", v, ..., "op", v)
# et      SUMIFS(TFTN[...],    TFTN[Compte],    "op", v, ..., "op", v)
SUMIFS_RE = re.compile(
    r'SUMIFS\((Balance|TFT)(N1?)\[(BE_Debit|BE_Credit|Mvt_Debit|Mvt_Credit|BS_Debit|BS_Credit)\],'
    r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],'
    r'"(>=|<=|>|<)(\d+)",'
    r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],'
    r'"(>=|<=|>|<)(\d+)"\)'
)
SUM_RANGE_RE = re.compile(r'SUM\(([A-Z]{1,2}\d{1,4}):([A-Z]{1,2}\d{1,4})\)')
CROSS_SHEET_RE = re.compile(
    r"(?:'([^']+)'|\b([A-Za-z][A-Za-z0-9 ]*))!\$?([A-Z]{1,3})\$?(\d{1,4})"
)
BARE_CELL_RE = re.compile(r'\b([A-Z]{1,2}\d{1,4})\b')
IFERROR_RE = re.compile(r'^IFERROR\((.*),""\)$')


def account_table_code(compte):
    s = str(compte).strip()
    s = (s + "00000000")[:5]
    return int(s)


def _ops(lo_op, lo, hi_op, hi):
    cmp = {
        ">=": lambda x, b: x >= b,
        "<=": lambda x, b: x <= b,
        ">": lambda x, b: x > b,
        "<": lambda x, b: x < b,
    }
    return cmp[lo_op], cmp[hi_op]


def make_sumifs(balanceN, balanceN1, tftn=None, tftn1=None):
    tables = {
        "BalanceN": balanceN,
        "BalanceN1": balanceN1,
        "TFTN": tftn or [],
        "TFTN1": tftn1 or [],
    }

    def SUMIFS_(table_key, field, lo_op, lo, hi_op, hi):
        rows = tables[table_key]
        lo_cmp, hi_cmp = _ops(lo_op, lo, hi_op, hi)
        total = 0.0
        for r in rows:
            t = r["table"]
            if lo_cmp(t, lo) and hi_cmp(t, hi):
                total += r.get(field, 0) or 0
        return total
    return SUMIFS_


def find_matching_paren(s, start):
    """s[start] doit etre '('. Retourne l'index de la parenthese fermante correspondante."""
    depth = 0
    for i in range(start, len(s)):
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


def split_top_level(s):
    """Coupe s par les virgules de profondeur 0 (en ignorant celles a l'interieur de parentheses)."""
    parts = []
    depth = 0
    cur = []
    for ch in s:
        if ch == "(":
            depth += 1
            cur.append(ch)
        elif ch == ")":
            depth -= 1
            cur.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return parts


def transpile(formula, sheet):
    s = formula[1:].strip()

    # IF(condition, alors, sinon) nu, enveloppant toute la formule.
    if s.startswith("IF("):
        close = find_matching_paren(s, 2)
        if close == len(s) - 1:
            inner = s[3:-1]
            parts = split_top_level(inner)
            if len(parts) == 3:
                cond_expr, _ = transpile("=" + parts[0], sheet)
                then_expr, _ = transpile("=" + parts[1], sheet)
                else_expr, _ = transpile("=" + parts[2], sheet)
                combined = "((%s) if (%s) else (%s))" % (then_expr, cond_expr, else_expr)
                return combined, False

    stash = []

    def put(snippet):
        stash.append(snippet)
        return "@@%d@@" % (len(stash) - 1)

    def repl_sumifs(m):
        base, suf, field, lo_op, lo, hi_op, hi = m.groups()
        table_key = base + suf
        return put("SUMIFS_('%s','%s','%s',%s,'%s',%s)" % (table_key, field, lo_op, lo, hi_op, hi))
    s = SUMIFS_RE.sub(repl_sumifs, s)

    def repl_sum(m):
        c1, c2 = m.groups()
        return put("SUMRANGE_('%s','%s','%s')" % (sheet, c1, c2))
    s = SUM_RANGE_RE.sub(repl_sum, s)

    def repl_cross(m):
        sh = m.group(1) or m.group(2)
        coord = m.group(3) + m.group(4)
        return put("GG('%s','%s')" % (sh, coord))
    s = CROSS_SHEET_RE.sub(repl_cross, s)

    m = IFERROR_RE.match(s)
    wrap_iferror = False
    if m:
        s = m.group(1)
        wrap_iferror = True

    s = BARE_CELL_RE.sub(lambda m: "G('%s')" % m.group(1), s)

    def restore(m):
        return stash[int(m.group(1))]
    while "@@" in s:
        s = re.sub(r'@@(\d+)@@', restore, s)

    s = s.replace("\n", " ")
    return s, wrap_iferror


def compute_sheet(sheet, balanceN, balanceN1, tftn=None, tftn1=None):
    cells = SHEETS_RAW[sheet]
    cache = {}
    SUMIFS_ = make_sumifs(balanceN, balanceN1, tftn, tftn1)

    def GG(sh, c):
        return 0

    def G(coord):
        if coord in cache:
            return cache[coord]
        raw = cells.get(coord)
        if raw is None:
            cache[coord] = 0
            return 0
        if not (isinstance(raw, str) and raw.startswith("=")):
            cache[coord] = raw if isinstance(raw, (int, float)) else 0
            return cache[coord]
        try:
            expr, wrap_iferror = transpile(raw, sheet)
            ns = {"SUMIFS_": SUMIFS_, "SUMRANGE_": SUMRANGE_, "G": G, "GG": GG, "ABS": abs, "ISBLANK": lambda v: False}
            if wrap_iferror:
                try:
                    val = eval(expr, ns)
                except Exception:
                    val = ""
            else:
                val = eval(expr, ns)
        except Exception:
            val = None
        cache[coord] = val
        return val

    def SUMRANGE_(sh, c1, c2):
        col1 = re.match(r'([A-Z]+)(\d+)', c1)
        col2 = re.match(r'([A-Z]+)(\d+)', c2)
        r1, r2 = int(col1.group(2)), int(col2.group(2))
        cstart, cend = ord(col1.group(1)), ord(col2.group(1))
        total = 0.0
        for co in range(cstart, cend + 1):
            for r in range(r1, r2 + 1):
                v = G("%s%d" % (chr(co), r))
                if isinstance(v, (int, float)):
                    total += v
        return total

    for coord in list(cells.keys()):
        row = int(re.search(r'\d+', coord).group())
        if 8 <= row <= 60:
            G(coord)
    return cache


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    if s in ("", "-", "–", "—"):
        return 0.0
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def tft_named_split(tftn, tftn1):
    """Reproduit les 8 plages nommees ECEN/ECN/ECHN/ECH/EPN/EPA/EPH/EPB de la
    feuille 'Table TFT' du classeur d'origine, utilisees par NOTE 34
    (analyse de la structure financiere) : pour chaque compte d'ecart de
    conversion (47811, 47812, 47911, 47912), la feuille calcule
    BE_Debit-BE_Credit+Mvt_Debit-Mvt_Credit puis scinde le resultat en une
    partie positive (colonnes G/P, table TFTN/TFTN1) et une partie negative
    inversee (colonnes H/Q), exactement comme la formule
    IF((...)>0, ..., 0) / IF((...)>0, 0, -(...)) repetee sur chaque ligne de
    detail. ECEN/ECHN/ECN/ECH utilisent la partie positive (G/P), EPN/EPH/
    EPA/EPB la partie negative (H/Q)."""
    def parts(rows, compte):
        code = account_table_code(compte)
        total = 0.0
        for r in rows or []:
            if r.get("table") == code:
                total += (
                    (r.get("BE_Debit") or 0) - (r.get("BE_Credit") or 0)
                    + (r.get("Mvt_Debit") or 0) - (r.get("Mvt_Credit") or 0)
                )
        pos = total if total > 0 else 0.0
        neg = -total if total < 0 else 0.0
        return pos, neg

    ecen, _ = parts(tftn, 47811)
    ecn, _ = parts(tftn1, 47811)
    echn, _ = parts(tftn, 47812)
    ech, _ = parts(tftn1, 47812)
    _, epn = parts(tftn, 47911)
    _, epa = parts(tftn1, 47911)
    _, eph = parts(tftn, 47912)
    _, epb = parts(tftn1, 47912)
    return {
        "ECEN": ecen, "ECN": ecn, "ECHN": echn, "ECH": ech,
        "EPN": epn, "EPA": epa, "EPH": eph, "EPB": epb,
    }


def build_balance_rows(entries):
    rows = []
    for e in entries:
        be_debit = _num(e.get("be_debit"))
        be_credit = _num(e.get("be_credit"))
        mvt_debit = _num(e.get("mvt_debit"))
        mvt_credit = _num(e.get("mvt_credit"))
        bs_debit = _num(e.get("bs_debit"))
        bs_credit = _num(e.get("bs_credit"))
        rows.append({
            "compte": e.get("compte"),
            "be_debit": be_debit,
            "be_credit": be_credit,
            "mvt_debit": mvt_debit,
            "mvt_credit": mvt_credit,
            "bs_debit": bs_debit,
            "bs_credit": bs_credit,
            "table": account_table_code(e.get("compte")),
            "BE_Debit": be_debit,
            "BE_Credit": be_credit,
            "Mvt_Debit": mvt_debit,
            "Mvt_Credit": mvt_credit,
            "BS_Debit": bs_debit,
            "BS_Credit": bs_credit,
        })
    return rows


# ---------------------------------------------------------------------------
# Moteur multi-feuilles pour la famille NOTE 3 (immobilisations) : permet la
# resolution de references inter-feuilles reelles (ex. NOTE 3D -> NOTE 3A/3C)
# et la prise en compte de saisies manuelles (cellules sans formule dans le
# classeur d'origine : virements, reevaluations, prix de cession, etc.).
# ---------------------------------------------------------------------------

def compute_workbook(sheets, balanceN, balanceN1, manual=None, row_range=(8, 60)):
    """
    sheets   : liste des feuilles a calculer (ex. ["NOTE 3A", "NOTE 3C", "NOTE 3D"])
    manual   : { "NOTE 3A": {"F12": 0, "G12": 0, ...}, ... } - saisies manuelles
               qui remplacent toute cellule vide/sans formule du classeur.
    Retourne : { "NOTE 3A": {"D12": ..., ...}, ... }
    """
    manual = manual or {}
    SUMIFS_ = make_sumifs(balanceN, balanceN1)
    caches = {}

    def SUMRANGE_(sh, c1, c2):
        col1 = re.match(r'([A-Z]+)(\d+)', c1)
        col2 = re.match(r'([A-Z]+)(\d+)', c2)
        r1, r2 = int(col1.group(2)), int(col2.group(2))
        cstart, cend = ord(col1.group(1)), ord(col2.group(1))
        total = 0.0
        for co in range(cstart, cend + 1):
            for r in range(r1, r2 + 1):
                v = GG(sh, "%s%d" % (chr(co), r))
                if isinstance(v, (int, float)):
                    total += v
        return total

    def GG(sh, coord):
        cache = caches.setdefault(sh, {})
        if coord in cache:
            return cache[coord]
        man = manual.get(sh, {})
        if coord in man and man[coord] not in (None, ""):
            val = _num(man[coord])
            cache[coord] = val
            return val
        cells = SHEETS_RAW.get(sh, {})
        raw = cells.get(coord)
        if raw is None:
            cache[coord] = 0
            return 0
        if not (isinstance(raw, str) and raw.startswith("=")):
            val = raw if isinstance(raw, (int, float)) else 0
            cache[coord] = val
            return val
        cache[coord] = 0  # garde anti-recursion (pas de cycles attendus)
        try:
            expr, wrap_iferror = transpile(raw, sh)
            G = lambda c: GG(sh, c)
            ns = {"SUMIFS_": SUMIFS_, "SUMRANGE_": SUMRANGE_, "G": G, "GG": GG, "ABS": abs, "ISBLANK": lambda v: False}
            if wrap_iferror:
                try:
                    val = eval(expr, ns)
                except Exception:
                    val = ""
            else:
                val = eval(expr, ns)
        except Exception:
            val = None
        cache[coord] = val
        return val

    for sh in sheets:
        cells = SHEETS_RAW.get(sh, {})
        for coord in list(cells.keys()):
            mrow = re.search(r'\d+', coord)
            if mrow and row_range[0] <= int(mrow.group()) <= row_range[1]:
                GG(sh, coord)
        for coord in manual.get(sh, {}):
            GG(sh, coord)
    return caches
