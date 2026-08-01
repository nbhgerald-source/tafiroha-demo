"""
Module de revision assistee TAFIROHA.

Analyse une balance (et les etats calcules) et produit une liste d'anomalies
classees par famille et par severite. Voir Controles_Revision_TAFIROHA.md pour
le catalogue complet et la justification de chaque controle.

Familles :
    MAP  integrite du mapping vers les etats (parametrage de l'application)
    INT  integrite de la balance
    SEN  comptes de sens anormal
    ATT  comptes d'attente et de regularisation
    FIS  coherence fiscale
    SOC  coherence sociale et paie
    IMM  immobilisations et amortissements
    STK  stocks
    VAR  variations N / N-1
    ETA  coherence inter-etats
    CPL  completude avant depot

Le moteur ne connait ni la base de donnees ni le web : il recoit un contexte
et retourne une liste de dictionnaires. La persistance est geree par app.py.
"""
import json
import os
import re
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- referentiel des sens attendus, clef 4 chiffres -------------------------
#
# Ce fichier est indispensable : sans lui, aucun compte ne se resout et tous
# sont vus comme absents du plan comptable. Un repli silencieux transformerait
# donc un fichier manquant en centaines de fausses anomalies, sans le moindre
# diagnostic. On memorise l'erreur pour la signaler explicitement et desactiver
# les controles qui en dependent (INT-06 et toute la famille SEN).
SENS_PLAN_ERREUR = None
try:
    _chemin = os.path.join(BASE_DIR, "revision_sens.json")
    with open(_chemin, encoding="utf-8") as fh:
        SENS_PLAN = {int(k): v for k, v in json.load(fh).items()}
    if not SENS_PLAN:
        SENS_PLAN_ERREUR = "le fichier revision_sens.json est vide"
except Exception as _e:
    SENS_PLAN = {}
    SENS_PLAN_ERREUR = "%s (%s)" % (_e, os.path.join(BASE_DIR, "revision_sens.json"))

# --- comptes supprimes ou remplaces par le SYSCOHADA revise ----------------
try:
    with open(os.path.join(BASE_DIR, "revision_obsoletes.json"), encoding="utf-8") as fh:
        OBSOLETES = {k: v for k, v in json.load(fh).items() if not k.startswith("_")}
except Exception:
    OBSOLETES = {}

BLOQUANT, MAJEUR, JUSTIFIER, INFO = "bloquant", "majeur", "justifier", "informatif"

SEVERITE_ORDRE = {BLOQUANT: 0, MAJEUR: 1, JUSTIFIER: 2, INFO: 3}
SEVERITE_LABEL = {
    BLOQUANT: "Bloquant",
    MAJEUR: "Majeur",
    JUSTIFIER: "A justifier",
    INFO: "Informatif",
}
FAMILLE_LABEL = {
    "MAP": "Integrite du mapping",
    "INT": "Integrite de la balance",
    "SEN": "Comptes de sens anormal",
    "ATT": "Comptes d'attente et de regularisation",
    "FIS": "Coherence fiscale",
    "SOC": "Coherence sociale et paie",
    "IMM": "Immobilisations et amortissements",
    "STK": "Stocks",
    "VAR": "Variations N / N-1",
    "ETA": "Coherence inter-etats",
    "CPL": "Completude avant depot",
}

# --- parametres par defaut, surchargeables par client ----------------------
DEFAUTS = {
    "seuil_absolu": 100000.0,      # montant en dessous duquel on n'alerte pas
    "seuil_variation_pct": 30.0,   # variation N/N-1 declenchant VAR-01
    "seuil_signification_pct": 1.0,  # % du total bilan servant de seuil
    "taux_tva": 18.0,              # taux normal de TVA
    "tva_tolerance_points": 3.0,   # ecart tolere sur le taux implicite
    "ratio_social_min": 10.0,      # 664 / 661 minimum attendu, en %
    "ratio_social_max": 40.0,      # maximum attendu, en %
}

# Nombre maximum d'occurrences detaillees par code avant condensation.
PLAFOND_DEFAUT = 40
PLAFOND = {
    "INT-08": 0,    # comptes sans libelle : purement indicatif
    "INT-10": 0,    # comptes de classe 9
    "MAP-09": 25,   # resolution au niveau 2 chiffres : le detail est diagnostique
    "INT-06": 15,   # comptes absents du plan
    "VAR-01": 20,   # variations significatives
    "VAR-02": 10,
    "VAR-03": 10,
    "ATT-08": 10,
    "FIS-08": 10,
    # Lignes de detail par compte : elles accompagnent une anomalie principale,
    # on en garde suffisamment pour couvrir les cas courants.
    "SOC-02D": 20,
    "SOC-03D": 20,
    "SOC-04D": 20,
    "IMM-02D": 20,
}

# En deca de ce nombre d'occurrences, le detail par compte est TOUJOURS affiche,
# quel que soit le plafond. Condenser trois lignes en une ligne de synthese fait
# perdre l'information sans rien gagner en lisibilite.
CONDENSATION_MINI = 12


# ---------------------------------------------------------------------------
# utilitaires
# ---------------------------------------------------------------------------
def _n(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def _key4(table):
    """Clef 4 chiffres utilisee pour la resolution du sens."""
    return int(str(int(table))[:4])


def resolve_sens(table, overrides=None, mappings=None):
    """Resolution par prefixe le plus long : 4 puis 3 chiffres.

    La resolution s'arrete a 3 chiffres. Au-dela, le rattachement deviendrait
    trop approximatif : deux comptes partageant seulement leurs deux premiers
    chiffres peuvent avoir des sens opposes (421 debiteur / 422 crediteur).
    Un compte non resolu doit etre rattache explicitement par l'utilisateur
    via un compte de regularisation, plutot que devine.

    Ordre de priorite :
      1. rattachement saisi par l'utilisateur (mappings)
      2. surcharge client (overrides)
      3. referentiel, prefixe 4 puis 3

    Retourne (sens, niveau). niveau vaut 4, 3 ou 0 si non resolu.
    """
    key = _key4(table)
    if mappings and key in mappings:
        key = mappings[key]
    if overrides and key in overrides:
        return overrides[key], 4
    s = str(key)
    for n in (4, 3):
        prefixe = s[:n]
        for k, v in SENS_PLAN.items():
            if str(k)[:n] == prefixe and v:
                return v, n
    return None, 0


def compte_obsolete(table):
    """Retourne (remplacant, libelle, severite) si le compte est supprime."""
    s = str(_key4(table))
    for n in (4, 3, 2):
        info = OBSOLETES.get(s[:n])
        if isinstance(info, dict):
            return (info.get("remplacant") or "",
                    info.get("libelle") or "Compte supprime du plan SYSCOHADA",
                    info.get("severite") or "majeur")
    return None


def _solde(r, quand="bs"):
    if quand == "be":
        return _n(r.get("be_debit")) - _n(r.get("be_credit"))
    if quand == "mvt":
        return _n(r.get("mvt_debit")) - _n(r.get("mvt_credit"))
    return _n(r.get("bs_debit")) - _n(r.get("bs_credit"))


def _somme(rows, lo, hi, champ="solde", quand="bs"):
    """Somme sur une plage de comptes [lo, hi[ exprimee en clef Table."""
    total = 0.0
    for r in rows or []:
        if lo <= r.get("table", 0) < hi:
            total += _solde(r, quand) if champ == "solde" else _n(r.get(champ))
    return total


def _dans(rows, lo, hi):
    return [r for r in (rows or []) if lo <= r.get("table", 0) < hi]


def _index(rows):
    """Indexe une balance par numero de compte."""
    return {str(r.get("compte")): r for r in (rows or [])}


class _Collecteur(object):
    def __init__(self, params):
        self.items = []
        self.p = params

    def add(self, code, famille, libelle, severite, compte=None, montant=None):
        self.items.append({
            "code": code,
            "famille": famille,
            "libelle": libelle,
            "severite": severite,
            "compte": str(compte) if compte is not None else None,
            "montant": round(_n(montant), 2) if montant is not None else None,
        })

    def seuil(self, montant):
        """Filtre de bruit : True si le montant merite une alerte."""
        return abs(_n(montant)) >= self.p["seuil_absolu"]


# ---------------------------------------------------------------------------
# MAP  integrite du mapping
# ---------------------------------------------------------------------------
_SUMIFS_RE = re.compile(
    r'SUMIFS\((?:Balance|TFT)(?:N1?)\[\w+\],'
    r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],"(>=|<=|>|<)(\d+)",'
    r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],"(>=|<=|>|<)(\d+)"\)'
)


_SUMIFS_TXT = (r'SUMIFS\((?:Balance|TFT)(?:N1?)\[\w+\],'
               r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],">=\d+",'
               r'(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],"<\d+"\)')
_PLATE_RE = re.compile(r'^=[+-]?' + _SUMIFS_TXT + r'(?:[+-]' + _SUMIFS_TXT + r')*$')
_TERME_RE = re.compile(r'([+-]?)(' + _SUMIFS_TXT + r')')
_DETAIL_RE = re.compile(r'\[(\w+)\],(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],'
                        r'">=(\d+)",(?:Balance|TFT)(?:N1?)\[(?:Table|Compte)\],"<(\d+)"')
_TABLE_RE = re.compile(r'SUMIFS\(((?:Balance|TFT)(?:N1?))\[')


def controles_map(c, sheets_raw, balN, balN1):
    plages = []
    for feuille, cellules in (sheets_raw or {}).items():
        if not isinstance(cellules, dict):
            continue
        for coord, formule in cellules.items():
            if not (isinstance(formule, str) and "SUMIFS" in formule):
                continue
            for m in _SUMIFS_RE.finditer(formule):
                lo_op, lo, hi_op, hi = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
                if lo > 99999 or hi > 99999:
                    c.add("MAP-02", "MAP",
                          "%s!%s : borne hors domaine (la clef Table plafonne a 5 chiffres) "
                          "- la formule retourne toujours zero" % (feuille, coord), BLOQUANT)
                elif lo_op.startswith(">") and hi_op.startswith("<") and lo >= hi:
                    c.add("MAP-01", "MAP",
                          "%s!%s : plage impossible [%d, %d[ - la formule retourne toujours zero"
                          % (feuille, coord, lo, hi), BLOQUANT)
                else:
                    plages.append((lo_op, lo, hi_op, hi))
    # MAP-04 : double captage a l'interieur d'une meme cellule.
    #
    # Limite volontaire : l'analyse ne porte que sur les formules qui sont une
    # somme PLATE de SUMIFS. Sur les formules contenant IF(...) ou des groupes
    # parenthetiques du type -(A+B), le signe effectif de chaque terme ne peut
    # pas etre determine par une simple lecture lineaire, et toute detection y
    # serait un faux positif. Mieux vaut ne rien dire que se tromper.
    for feuille, cellules in (sheets_raw or {}).items():
        if not isinstance(cellules, dict):
            continue
        for coord, formule in cellules.items():
            if not (isinstance(formule, str) and "SUMIFS" in formule):
                continue
            compact = "".join(formule.split())
            if not _PLATE_RE.match(compact):
                continue
            termes = []
            for m in _TERME_RE.finditer(compact):
                signe = m.group(1) or "+"
                seg = m.group(2)
                det = _DETAIL_RE.search(seg)
                tab = _TABLE_RE.search(seg)
                if not (det and tab):
                    continue
                champ, lo, hi = det.groups()
                termes.append((signe, tab.group(1), champ, int(lo), int(hi)))
            for i, (s, tb, ch, a, b) in enumerate(termes):
                for (s2, tb2, ch2, c2, e2) in termes[i + 1:]:
                    if (s, tb, ch) != (s2, tb2, ch2):
                        continue
                    lo, hi = max(a, c2), min(b, e2)
                    if lo < hi:
                        c.add("MAP-04", "MAP",
                              "%s!%s : les comptes %d a %d sont captes deux fois par "
                              "les plages [%d, %d[ et [%d, %d[ - le montant est double"
                              % (feuille, coord, lo, hi - 1, a, b, c2, e2), MAJEUR)

    if not plages:
        return
    cmp_ = {">=": lambda x, b: x >= b, "<=": lambda x, b: x <= b,
            ">": lambda x, b: x > b, "<": lambda x, b: x < b}

    def capte(t):
        return any(cmp_[a](t, lo) and cmp_[b](t, hi) for a, lo, b, hi in plages)

    for label, rows in (("N", balN), ("N-1", balN1)):
        for r in rows or []:
            if not capte(r.get("table", 0)):
                c.add("MAP-03", "MAP",
                      "Compte present en balance %s mais capte par aucun etat : "
                      "il disparait du bilan" % label, BLOQUANT,
                      r.get("compte"), _solde(r))


# ---------------------------------------------------------------------------
# INT  integrite de la balance
# ---------------------------------------------------------------------------
def controles_int(c, balN, balN1, overrides=None, mappings=None):
    for label, rows in (("N", balN), ("N-1", balN1)):
        if not rows:
            continue
        for code, champ_d, champ_c, nom in (
            ("INT-02", "be_debit", "be_credit", "balance d'entree"),
            ("INT-01", "mvt_debit", "mvt_credit", "mouvements"),
            ("INT-03", "bs_debit", "bs_credit", "balance de sortie"),
        ):
            d = sum(_n(r.get(champ_d)) for r in rows)
            cr = sum(_n(r.get(champ_c)) for r in rows)
            if abs(d - cr) > 0.5:
                c.add(code, "INT",
                      "Balance %s desequilibree sur les %s : debit %s / credit %s"
                      % (label, nom, _fmt(d), _fmt(cr)), BLOQUANT, None, d - cr)

        vus = {}
        for r in rows:
            cpt = str(r.get("compte"))
            ecart = _solde(r, "be") + _solde(r, "mvt") - _solde(r)
            if abs(ecart) > 0.5:
                c.add("INT-04", "INT",
                      "Balance %s : incoherence arithmetique, "
                      "balance d'entree + mouvements different du solde de sortie" % label,
                      BLOQUANT, cpt, ecart)
            if cpt in vus:
                c.add("INT-05", "INT",
                      "Balance %s : compte present en double" % label, BLOQUANT, cpt)
            vus[cpt] = True
            if not str(r.get("designation") or "").strip():
                c.add("INT-08", "INT",
                      "Balance %s : compte sans libelle" % label, INFO, cpt)
            # compte supprime du referentiel SYSCOHADA revise
            obs = compte_obsolete(r.get("table", 0))
            if obs:
                remplacant, libelle, sev = obs
                # un rattachement deja enregistre route correctement le montant :
                # l'alerte se transforme en simple rappel de renumerotation
                rattache = bool(mappings and _key4(r.get("table", 0)) in mappings)
                if rattache:
                    msg = ("Balance %s : compte supprime du SYSCOHADA, deja rattache "
                           "au compte %s. Pensez a le renumeroter dans votre "
                           "comptabilite." % (label, mappings[_key4(r.get("table", 0))]))
                    severite = INFO
                else:
                    msg = "Balance %s : %s" % (label, libelle)
                    if remplacant:
                        msg += ". Utiliser desormais le compte %s" % remplacant
                    severite = BLOQUANT if sev == "bloquant" else MAJEUR
                c.add("INT-13", "INT", msg, severite, cpt, _solde(r))

            # sans referentiel charge, la resolution est impossible : on ne
            # signale rien plutot que de declarer toute la balance hors plan
            if SENS_PLAN and not obs:
                sens, niveau = resolve_sens(r.get("table", 0), overrides, mappings)
                # La resolution a 3 chiffres est fiable : les sous-comptes d'une
                # meme racine a 3 chiffres partagent leur sens (4011 a 4017 tous
                # crediteurs). Seul l'echec complet merite un signalement.
                if niveau == 0:
                    c.add("INT-06", "INT",
                          "Balance %s : compte non rattachable au plan SYSCOHADA. "
                          "Indiquez un compte de regularisation : le rattachement sera "
                          "conserve pour cet exercice et les suivants." % label,
                          MAJEUR, cpt, _solde(r))
            if r.get("table", 0) >= 90000:
                c.add("INT-10", "INT",
                      "Compte de classe 9 present dans la balance generale", INFO, cpt)

    # INT-12 : report du resultat N-1 en balance d'entree N.
    #
    # A la cloture de N-1 le resultat vit encore dans les classes 6 a 8 ; il est
    # vire en 12xx/13xx a l'ouverture de N. Le solde d'ouverture de ces comptes
    # ne peut donc PAS etre compare a leur solde de cloture N-1 : l'ecart normal
    # vaut exactement le resultat de l'exercice precedent. Le classeur d'origine
    # applique le meme ajustement dans son controle global (TABLE DE CONVERSION
    # M6/Z6, qui corrige les totaux par les classes 6 a 9).
    #
    # Le numero de compte porteur du resultat change souvent d'un exercice a
    # l'autre (139100 en N-1, 130900 en N sur un dossier reel) : le controle se
    # fait donc sur l'agregat 12xx+13xx, jamais compte par compte.
    if balN and balN1:
        resultat_n1 = -_somme(balN1, 60000, 90000)
        be_n = _somme(balN, 12000, 14000, "solde", "be")
        bs_n1 = _somme(balN1, 12000, 14000)
        attendu = bs_n1 - resultat_n1
        ecart = be_n - attendu
        if abs(ecart) > 0.5:
            c.add("INT-12", "INT",
                  "Report du resultat N-1 incorrect : la balance d'entree des comptes "
                  "12xx et 13xx est de %s, alors que la cloture N-1 (%s) augmentee du "
                  "resultat N-1 (%s) donne %s"
                  % (_fmt(be_n), _fmt(bs_n1), _fmt(-resultat_n1), _fmt(attendu)),
                  BLOQUANT, None, ecart)

    # INT-09 : rupture de report a nouveau, compte par compte
    if balN and balN1:
        iN, i1 = _index(balN), _index(balN1)
        ecarts = []
        for cpt in set(iN) | set(i1):
            rN, r1 = iN.get(cpt), i1.get(cpt)
            table = (rN or r1).get("table", 0)
            if table >= 60000:
                continue
            # comptes de report a nouveau et de resultat : traites par INT-12
            if 12000 <= table < 14000:
                continue
            be = _solde(rN, "be") if rN else 0.0
            bs = _solde(r1) if r1 else 0.0
            if abs(be - bs) > 0.5:
                ecarts.append((cpt, be - bs))
        # un reclassement se compense : deux ecarts opposes -> informatif
        restants = list(ecarts)
        apparies = set()
        for i, (c1, e1) in enumerate(ecarts):
            if c1 in apparies:
                continue
            for c2, e2 in ecarts[i + 1:]:
                if c2 in apparies:
                    continue
                if abs(e1 + e2) < 0.5:
                    apparies.add(c1)
                    apparies.add(c2)
                    c.add("INT-09", "INT",
                          "Reclassement entre exercices : le solde d'ouverture N "
                          "differe de la cloture N-1, compense par le compte %s" % c2,
                          INFO, c1, e1)
                    c.add("INT-09", "INT",
                          "Reclassement entre exercices : le solde d'ouverture N "
                          "differe de la cloture N-1, compense par le compte %s" % c1,
                          INFO, c2, e2)
                    break
        for cpt, ec in restants:
            if cpt in apparies:
                continue
            c.add("INT-09", "INT",
                  "Rupture de report a nouveau : le solde d'ouverture N ne correspond "
                  "pas au solde de cloture N-1", BLOQUANT, cpt, ec)


def _fmt(x):
    return "{:,.0f}".format(_n(x)).replace(",", " ")


# ---------------------------------------------------------------------------
# SEN  comptes de sens anormal
# ---------------------------------------------------------------------------
def controles_sen(c, balN, overrides=None, mappings=None):
    if not SENS_PLAN:
        return
    for r in balN or []:
        table = r.get("table", 0)
        sens, niveau = resolve_sens(table, overrides, mappings)
        solde = _solde(r)
        cpt = r.get("compte")

        # "N" : compte de transit qui doit etre solde a la cloture. Ce n'est pas
        # un sens mais une exigence de solde nul (virements de fonds, virements
        # internes) : tout solde residuel signale une operation non affectee.
        if sens == "N":
            if abs(solde) >= 0.5:
                c.add("SEN-18", "SEN",
                      "Compte de transit non solde a la cloture : "
                      "une operation n'a pas ete affectee", BLOQUANT, cpt, solde)
            continue

        # Comptes bancaires : signalement independant du sens attendu. Ces
        # comptes sont parametres en "M" car un solde crediteur n'est pas une
        # erreur, mais il merite d'etre porte a la connaissance du reviseur.
        if 52000 <= table < 52700 and solde < -0.5:
            if c.seuil(solde):
                c.add("SEN-13", "SEN",
                      "Solde bancaire crediteur : a verifier", JUSTIFIER, cpt, solde)
            continue

        if sens in (None, "M") or niveau == 0:
            continue
        if abs(solde) < 0.5:
            continue
        anormal = (sens == "D" and solde < 0) or (sens == "C" and solde > 0)
        if not anormal:
            continue

        # cas particuliers, plus parlants que le message generique
        if 57000 <= table < 58000:
            c.add("SEN-01", "SEN",
                  "Caisse creditrice : physiquement impossible", BLOQUANT, cpt, solde)
        elif 31000 <= table < 39000:
            c.add("SEN-12", "SEN",
                  "Stock negatif", BLOQUANT, cpt, solde)
        elif 28000 <= table < 29000:
            c.add("SEN-09", "SEN",
                  "Amortissements cumules au debit : erreur de sens sur une dotation "
                  "ou une sortie", BLOQUANT, cpt, solde)
        elif 10100 <= table < 10500:
            c.add("SEN-08", "SEN",
                  "Capital au debit : erreur d'imputation", BLOQUANT, cpt, solde)
        elif not c.seuil(solde):
            continue
        else:
            attendu = "debiteur" if sens == "D" else "crediteur"
            constate = "crediteur" if sens == "D" else "debiteur"
            severite = MAJEUR if niveau >= 3 else JUSTIFIER
            c.add("SEN-00", "SEN",
                  "Sens anormal : compte attendu %s, constate %s" % (attendu, constate),
                  severite, cpt, solde)


# ---------------------------------------------------------------------------
# ATT  comptes d'attente et de regularisation
# ---------------------------------------------------------------------------
def controles_att(c, balN):
    REGLES = [
        ("ATT-01", 47500, 47600, BLOQUANT,
         "Compte transitoire d'ajustement non solde : une operation n'a pas ete affectee"),
        ("ATT-02", 58500, 58600, BLOQUANT,
         "Compte de virements internes non solde"),
        ("ATT-03", 18500, 18700, MAJEUR,
         "Compte de liaison non solde"),
        ("ATT-04", 47600, 47700, JUSTIFIER,
         "Charges constatees d'avance : verifier le calcul du prorata"),
        ("ATT-05", 47700, 47800, JUSTIFIER,
         "Produits constates d'avance : verifier le calcul du prorata"),
        ("ATT-06", 47800, 47900, MAJEUR,
         "Ecart de conversion Actif : verifier la provision pour perte de change"),
        ("ATT-07", 47900, 48000, JUSTIFIER,
         "Ecart de conversion Passif a justifier"),
    ]
    for code, lo, hi, sev, msg in REGLES:
        for r in _dans(balN, lo, hi):
            solde = _solde(r)
            if abs(solde) < 0.5:
                continue
            if sev in (JUSTIFIER, MAJEUR) and not c.seuil(solde):
                continue
            c.add(code, "ATT", msg, sev, r.get("compte"), solde)

    # ATT-08 s'appuie sur le seuil de signification, pas sur le seuil absolu :
    # sans cela le controle remonte des dizaines de lignes sur les gros dossiers.
    total = abs(_somme(balN, 10000, 60000))
    seuil = max(c.p["seuil_absolu"] * 10,
                total * c.p["seuil_signification_pct"] / 100.0)
    for r in _dans(balN, 47100, 47500):
        solde = _solde(r)
        if abs(solde) < seuil:
            continue
        # Un solde conforme au sens attendu du compte n'a rien d'anormal :
        # 4712 "Crediteurs divers" est crediteur par nature. Seuls les comptes
        # dont le sens est indetermine, ou dont le solde part dans le mauvais
        # sens, meritent une justification. Les seconds sont deja couverts par
        # la famille SEN, on ne les redouble pas ici.
        sens, niveau = resolve_sens(r.get("table", 0))
        if sens in ("D", "C") and niveau > 0:
            continue
        c.add("ATT-08", "ATT",
              "Compte divers d'un montant significatif, au sens indetermine : "
              "a justifier", JUSTIFIER, r.get("compte"), solde)


# ---------------------------------------------------------------------------
# FIS  coherence fiscale
# ---------------------------------------------------------------------------
def controles_fis(c, balN, resultat):
    p = c.p
    # TVA facturee de l'exercice = mouvements CREDITEURS des comptes 443x. Le
    # debit de ces comptes correspond au virement vers la TVA due (4441) en fin
    # de periode declarative : le retenir annulerait la mesure.
    tva_col = abs(_somme(balN, 44300, 44400, "mvt_credit"))
    # Chiffre d'affaires NET : credit moins debit, pour tenir compte des rabais,
    # remises et ristournes accordes (7019, 7029...) qui viennent en diminution.
    ca = _somme(balN, 70000, 71000, "mvt_credit") - _somme(balN, 70000, 71000, "mvt_debit")
    ca = abs(ca)
    tva_rec = abs(_somme(balN, 44500, 44600, "mvt_debit"))
    achats = abs(_somme(balN, 60000, 61000, "mvt_debit"))

    if ca > 0 and tva_col > 0:
        taux = 100.0 * tva_col / ca
        if abs(taux - p["taux_tva"]) > p["tva_tolerance_points"]:
            # On expose le detail du calcul : sans les composantes, le taux seul
            # ne permet aucune analyse. Le CA taxable implicite indique quelle
            # part du chiffre d'affaires porte effectivement de la TVA.
            ca_taxable = tva_col / (p["taux_tva"] / 100.0) if p["taux_tva"] else 0
            part = 100.0 * ca_taxable / ca if ca else 0
            c.add("FIS-01", "FIS",
                  "Taux de TVA collectee implicite de %.2f %% contre %.1f %% attendu. "
                  "Calcul : TVA facturee %s (comptes 443x, mouvements crediteurs) "
                  "rapportee au chiffre d'affaires net %s (comptes 70x). "
                  "Au taux de %.1f %%, cette TVA correspondrait a un chiffre d'affaires "
                  "taxable de %s, soit %.1f %% du total : verifier que le solde releve "
                  "bien d'operations exonerees ou exportees."
                  % (taux, p["taux_tva"], _fmt(tva_col), _fmt(ca),
                     p["taux_tva"], _fmt(ca_taxable), part),
                  JUSTIFIER, None, tva_col)
    if tva_col > 0 and ca == 0:
        c.add("FIS-02", "FIS",
              "TVA collectee sans chiffre d'affaires comptabilise", MAJEUR, None, tva_col)
    if ca > 0 and tva_col == 0:
        c.add("FIS-03", "FIS",
              "Chiffre d'affaires sans TVA collectee : verifier le regime d'exoneration",
              JUSTIFIER, None, ca)
    if tva_rec > 0 and achats == 0:
        c.add("FIS-04", "FIS",
              "TVA recuperable de %s comptabilisee (comptes 445x) sans aucun achat "
              "sur les comptes 60x" % _fmt(tva_rec), MAJEUR, None, tva_rec)
    elif tva_rec > 0 and achats > 0:
        taux_rec = 100.0 * tva_rec / achats
        if abs(taux_rec - p["taux_tva"]) > p["tva_tolerance_points"]:
            c.add("FIS-14", "FIS",
                  "Taux de TVA recuperable implicite de %.2f %% contre %.1f %% attendu. "
                  "Calcul : TVA recuperable %s (comptes 445x) rapportee aux achats %s "
                  "(comptes 60x). Verifier les achats non soumis a TVA et la TVA non "
                  "deductible." % (taux_rec, p["taux_tva"], _fmt(tva_rec), _fmt(achats)),
                  JUSTIFIER, None, tva_rec)

    tva_immo = abs(_somme(balN, 44510, 44520, "mvt_debit"))
    acq_immo = _somme(balN, 21000, 28000, "mvt_debit")
    if tva_immo > 0 and acq_immo == 0:
        c.add("FIS-05", "FIS",
              "TVA recuperable sur immobilisations sans acquisition d'immobilisation",
              MAJEUR, None, tva_immo)

    for r in _dans(balN, 44700, 44800):
        solde = _solde(r)
        if c.seuil(solde):
            c.add("FIS-08", "FIS",
                  "Retenues a la source non reversees a la cloture : solde de %s sur le "
                  "compte %s" % (_fmt(solde), r.get("compte")),
                  JUSTIFIER, r.get("compte"), solde)

    impot = abs(_somme(balN, 89000, 90000, "mvt_debit"))
    resultat_av = _somme(balN, 70000, 80000, "mvt_credit") - _somme(balN, 60000, 70000, "mvt_debit")
    if impot == 0 and resultat_av > 0:
        c.add("FIS-09", "FIS",
              "Aucun impot sur le resultat comptabilise alors que le resultat "
              "avant impot est beneficiaire", MAJEUR, None, resultat_av)
    if impot == 0 and resultat_av <= 0:
        c.add("FIS-12", "FIS",
              "Resultat deficitaire : verifier l'application de l'impot minimum forfaitaire",
              JUSTIFIER, None, resultat_av)


# ---------------------------------------------------------------------------
# SOC  coherence sociale et paie
# ---------------------------------------------------------------------------
def _detail_comptes(c, code, famille, rows, lo, hi, champ, quand, libelle, mini=None):
    """Emet une ligne informative par compte contribuant a un agregat.

    Sans ce detail, un controle portant sur une plage de comptes affiche un
    montant sans dire d'ou il vient : le reviseur ne peut ni le verifier ni
    savoir ou intervenir."""
    seuil = mini if mini is not None else c.p["seuil_absolu"] / 10.0
    for r in _dans(rows, lo, hi):
        v = _solde(r, quand) if champ == "solde" else _n(r.get(champ))
        if abs(v) < seuil:
            continue
        c.add(code, famille, libelle, INFO, r.get("compte"), v)


def controles_soc(c, balN):
    p = c.p
    salaires = abs(_somme(balN, 66100, 66400, "mvt_debit"))
    charges = abs(_somme(balN, 66400, 66500, "mvt_debit"))

    if charges > 0 and salaires == 0:
        c.add("SOC-01", "SOC",
              "Charges sociales de %s comptabilisees (comptes 664x) sans aucune "
              "remuneration sur les comptes 661x a 663x" % _fmt(charges),
              MAJEUR, None, charges)

    if salaires > 0 and charges > 0:
        ratio = 100.0 * charges / salaires
        if ratio < p["ratio_social_min"] or ratio > p["ratio_social_max"]:
            c.add("SOC-02", "SOC",
                  "Ratio charges sociales sur salaires de %.1f %%, hors de l'intervalle "
                  "attendu (%.0f a %.0f %%). Calcul : charges sociales %s (comptes 664x) "
                  "rapportees aux remunerations %s (comptes 661x a 663x). Verifier les "
                  "exonerations, le personnel detache et les charges comptabilisees "
                  "ailleurs." % (ratio, p["ratio_social_min"], p["ratio_social_max"],
                                 _fmt(charges), _fmt(salaires)),
                  JUSTIFIER, None, charges)
            _detail_comptes(c, "SOC-02D", "SOC", balN, 66400, 66500, "mvt_debit",
                            None, "Charge sociale de l'exercice")
            _detail_comptes(c, "SOC-02D", "SOC", balN, 66100, 66400, "mvt_debit",
                            None, "Remuneration de l'exercice")

    if salaires > 0:
        du = abs(_somme(balN, 42200, 42300))
        mois = du / (salaires / 12.0) if salaires else 0
        if du > salaires / 12.0 * 2:
            c.add("SOC-03", "SOC",
                  "Remunerations dues de %s au bilan (comptes 422x), soit %.1f mois de "
                  "salaire pour une masse salariale annuelle de %s"
                  % (_fmt(du), mois, _fmt(salaires)), JUSTIFIER, None, du)
            _detail_comptes(c, "SOC-03D", "SOC", balN, 42200, 42300, "solde",
                            "bs", "Remuneration due")

    dette_soc = abs(_somme(balN, 43000, 44000))
    if salaires > 0 and dette_soc > salaires / 2:
        mois = dette_soc / (salaires / 12.0)
        c.add("SOC-04", "SOC",
              "Dette envers les organismes sociaux de %s (comptes 43x), soit %.1f mois "
              "de salaire pour une masse salariale annuelle de %s. Verifier les "
              "cotisations non reversees et les eventuels moratoires."
              % (_fmt(dette_soc), mois, _fmt(salaires)), JUSTIFIER, None, dette_soc)
        _detail_comptes(c, "SOC-04D", "SOC", balN, 43000, 44000, "solde",
                        "bs", "Dette sociale au bilan")


# ---------------------------------------------------------------------------
# IMM  immobilisations et amortissements
# ---------------------------------------------------------------------------
def controles_imm(c, balN, balN1, manual):
    # Correspondance SYSCOHADA immobilisation -> amortissement :
    # 21x -> 281x, 22x -> 282x, 23x -> 283x, 24x -> 284x. Ce n'est pas un
    # decalage constant, d'ou le tableau explicite.
    CORRESP = (
        (21000, 22000, 28100, 28200, "immobilisations incorporelles", True),
        (22000, 23000, 28200, 28300, "terrains", False),
        (23000, 24000, 28300, 28400, "batiments et installations", True),
        (24000, 25000, 28400, 28500, "materiel et mobilier", True),
    )
    for lo, hi, alo, ahi, nom, amortissable in CORRESP:
        brut = _somme(balN, lo, hi)
        amort = abs(_somme(balN, alo, ahi))
        if brut > 0 and amort > brut + 0.5:
            c.add("IMM-01", "IMM",
                  "Amortissements cumules (%s) superieurs a la valeur brute (%s) sur les %s"
                  % (_fmt(amort), _fmt(brut), nom), BLOQUANT, str(lo // 1000), amort - brut)
        if amortissable and brut > 0 and amort == 0:
            c.add("IMM-03", "IMM",
                  "Aucun amortissement comptabilise sur les %s alors que le brut est de %s"
                  % (nom, _fmt(brut)), MAJEUR, str(lo // 1000), brut)
        if amortissable and brut > 0 and abs(amort - brut) < 0.5:
            c.add("IMM-04", "IMM",
                  "Les %s sont totalement amorties : verifier si les biens sont "
                  "encore en service" % nom, INFO, str(lo // 1000), brut)

    # Dotation portee en charge (681x) contre dotation creditee aux comptes
    # d'amortissement (28xx). Ces deux montants doivent etre egaux : c'est la
    # meme ecriture vue de ses deux cotes.
    charge = abs(_somme(balN, 68100, 68200, "mvt_debit"))
    creditee = abs(_somme(balN, 28000, 29000, "mvt_credit"))
    sorties = abs(_somme(balN, 28000, 29000, "mvt_debit"))
    cumul_n = abs(_somme(balN, 28000, 29000))
    cumul_n1 = abs(_somme(balN1, 28000, 29000))
    variation = cumul_n - cumul_n1
    ecart = charge - creditee
    if (charge or creditee) and abs(ecart) > c.p["seuil_absolu"]:
        c.add("IMM-02", "IMM",
              "Dotation aux amortissements incoherente. Detail du calcul : "
              "dotation portee en charge (comptes 681x) %s, dotation creditee aux "
              "amortissements (comptes 28xx) %s, ecart %s. Pour memoire : cumul N %s, "
              "cumul N-1 %s, variation %s, sorties de l'exercice %s."
              % (_fmt(charge), _fmt(creditee), _fmt(ecart),
                 _fmt(cumul_n), _fmt(cumul_n1), _fmt(variation), _fmt(sorties)),
              MAJEUR, None, ecart)
        # Detail par categorie : indique ou se situe l'ecart.
        for lo, hi, nom in ((28100, 28200, "281 incorporelles"),
                            (28200, 28300, "282 terrains"),
                            (28300, 28400, "283 batiments"),
                            (28400, 28500, "284 materiel"),
                            (28500, 29000, "285-289 autres")):
            cn = abs(_somme(balN, lo, hi))
            c1 = abs(_somme(balN1, lo, hi))
            dot = abs(_somme(balN, lo, hi, "mvt_credit"))
            sor = abs(_somme(balN, lo, hi, "mvt_debit"))
            if not (cn or c1 or dot or sor):
                continue
            c.add("IMM-02D", "IMM",
                  "%s : cumul N %s, cumul N-1 %s, dotation de l'exercice %s, "
                  "sorties %s" % (nom, _fmt(cn), _fmt(c1), _fmt(dot), _fmt(sor)),
                  INFO, str(lo // 10), cn - c1)

    prod_cession = abs(_somme(balN, 82000, 83000, "mvt_credit"))
    if prod_cession > 0 and sorties == 0:
        c.add("IMM-05", "IMM",
              "Produit de cession comptabilise sans sortie d'amortissement",
              MAJEUR, None, prod_cession)

    # recoupement avec l'inventaire SUPPL4
    inv = (manual or {}).get("SUPPL4") or {}
    if inv:
        tot_acq = sum(_n(v) for k, v in inv.items() if str(k).upper().startswith("G"))
        brut_bal = _somme(balN, 21000, 28000)
        if tot_acq > 0 and abs(tot_acq - brut_bal) > c.p["seuil_absolu"]:
            c.add("IMM-08", "IMM",
                  "Ecart entre l'inventaire SUPPL4 (%s) et les comptes d'immobilisations "
                  "de la balance (%s)" % (_fmt(tot_acq), _fmt(brut_bal)),
                  MAJEUR, None, tot_acq - brut_bal)


# ---------------------------------------------------------------------------
# STK  stocks
# ---------------------------------------------------------------------------
def controles_stk(c, balN, balN1):
    # variation de stocks : mvt(603x) doit valoir -(SF - SI)
    for lo, hi, lo_v, hi_v, nom in (
        (31000, 32000, 60310, 60320, "marchandises"),
        (32000, 33000, 60320, 60330, "matieres premieres"),
        (33000, 34000, 60330, 60340, "autres approvisionnements"),
    ):
        sf = _somme(balN, lo, hi)
        si = _somme(balN, lo, hi, "solde", "be")
        var = _somme(balN, lo_v, hi_v, "solde", "mvt")
        if abs(sf - si) < 0.5 and abs(var) < 0.5:
            continue
        ecart = var + (sf - si)
        if abs(ecart) > c.p["seuil_absolu"]:
            c.add("STK-01", "STK",
                  "Variation de stock %s incoherente : ecriture de variation %s, "
                  "variation reelle du stock %s" % (nom, _fmt(var), _fmt(sf - si)),
                  BLOQUANT, None, ecart)
        if abs(sf - si) > 0.5 and abs(var) < 0.5:
            c.add("STK-03", "STK",
                  "Stock %s modifie sans ecriture de variation" % nom,
                  BLOQUANT, None, sf - si)

    for lo, hi in ((31000, 39000),):
        for r in _dans(balN, lo, hi):
            table = r.get("table", 0)
            if 39000 <= table:
                continue
    depr = abs(_somme(balN, 39000, 40000))
    brut = _somme(balN, 31000, 39000)
    if depr > brut + 0.5:
        c.add("STK-04", "STK",
              "Depreciation des stocks superieure a la valeur brute", BLOQUANT, None, depr - brut)

    sf_tot = _somme(balN, 31000, 39000)
    sf_n1 = _somme(balN1, 31000, 39000)
    if sf_tot > 0 and abs(sf_tot - sf_n1) < 0.5:
        c.add("STK-05", "STK",
              "Stock strictement identique a N-1 : l'inventaire physique a-t-il ete realise ?",
              MAJEUR, None, sf_tot)


# ---------------------------------------------------------------------------
# VAR  variations N / N-1
# ---------------------------------------------------------------------------
def controles_var(c, balN, balN1, bilan):
    p = c.p
    iN, i1 = _index(balN), _index(balN1)
    total_bilan = abs(_somme(balN, 10000, 60000))
    seuil = max(p["seuil_absolu"] * 10,
                total_bilan * p["seuil_signification_pct"] / 100.0)

    for cpt in set(iN) | set(i1):
        rN, r1 = iN.get(cpt), i1.get(cpt)
        sN = _solde(rN) if rN else 0.0
        s1 = _solde(r1) if r1 else 0.0
        ecart = sN - s1
        if rN and not r1 and abs(sN) >= seuil:
            c.add("VAR-03", "VAR", "Compte nouveau en N", INFO, cpt, sN)
            continue
        if r1 and not rN and abs(s1) >= seuil:
            c.add("VAR-02", "VAR",
                  "Compte mouvemente en N-1 et absent en N", JUSTIFIER, cpt, s1)
            continue
        if abs(ecart) >= seuil and abs(s1) > 0:
            pct = 100.0 * ecart / abs(s1)
            if abs(pct) >= p["seuil_variation_pct"]:
                c.add("VAR-01", "VAR",
                      "Variation de %+.0f %% par rapport a N-1 (%s -> %s)"
                      % (pct, _fmt(s1), _fmt(sN)), JUSTIFIER, cpt, ecart)

    cp = -_somme(balN, 10000, 20000)
    capital = -_somme(balN, 10100, 10500)
    if capital > 0 and cp < capital / 2:
        c.add("VAR-09", "VAR",
              "Capitaux propres inferieurs a la moitie du capital social : "
              "regularisation exigee par l'Acte uniforme OHADA", MAJEUR, None, cp)

    for r in _dans(balN, 46200, 46300):
        solde = _solde(r)
        if solde > 0 and c.seuil(solde):
            c.add("VAR-10", "VAR",
                  "Compte courant d'associe debiteur : interdit dans certaines formes sociales",
                  MAJEUR, r.get("compte"), solde)


# ---------------------------------------------------------------------------
# ETA  coherence inter-etats
# ---------------------------------------------------------------------------
def controles_eta(c, bilan, resultat, tft, balN, balN1=None, notes=None):
    def v(d, k):
        return _n((d or {}).get(k))

    # Coordonnees reelles du classeur : BILAN actif total general = H39,
    # passif total general = M39, RESULTAT net = I52, TFT tresorerie 31/12 = I40.
    tol = max(1.0, c.p["seuil_absolu"] / 1000.0)

    actif = v(bilan, "H39")
    passif = v(bilan, "M39")
    if actif and passif and abs(actif - passif) > tol:
        c.add("ETA-01", "ETA",
              "Bilan desequilibre : total actif %s, total passif %s" % (_fmt(actif), _fmt(passif)),
              BLOQUANT, None, actif - passif)

    # A la cloture de N le resultat n'est pas encore vire en 13xx : il vit dans
    # les classes 6 a 8. La reference est donc le solde de ces classes, pas le
    # compte 13. Ce controle valide le mapping de la feuille RESULTAT.
    res_cr = v(resultat, "I52")
    res_bal = -_somme(balN, 60000, 90000)
    if res_cr and abs(res_cr - res_bal) > tol:
        c.add("ETA-02", "ETA",
              "Resultat du compte de resultat (%s) different du resultat calcule "
              "directement sur les classes 6 a 8 de la balance (%s)"
              % (_fmt(res_cr), _fmt(res_bal)), BLOQUANT, None, res_cr - res_bal)

    # ── TFT : coherence interne et rattachement a la balance ────────────────
    A = v(tft, "I10")     # tresorerie au 1er janvier
    B = v(tft, "I19")     # flux operationnels
    C_ = v(tft, "I26")    # flux d'investissement
    F = v(tft, "I38")     # flux de financement (D+E)
    VARI = v(tft, "I39")  # variation de la periode
    FIN = v(tft, "I40")   # tresorerie au 31 decembre
    tres_n = _somme(balN, 50000, 60000)
    tres_n1 = _somme(balN1, 50000, 60000)

    if any((A, B, C_, F, VARI, FIN)):
        if abs(B + C_ + F - VARI) > tol:
            c.add("ETA-10", "ETA",
                  "TFT : la variation de tresorerie (%s) ne correspond pas a la somme "
                  "des flux operationnels, d'investissement et de financement (%s)"
                  % (_fmt(VARI), _fmt(B + C_ + F)), BLOQUANT, None, B + C_ + F - VARI)
        if abs(A + VARI - FIN) > tol:
            c.add("ETA-11", "ETA",
                  "TFT : tresorerie d'ouverture (%s) plus variation (%s) ne donne pas "
                  "la tresorerie de cloture (%s)" % (_fmt(A), _fmt(VARI), _fmt(FIN)),
                  BLOQUANT, None, A + VARI - FIN)
        if abs(FIN - tres_n) > tol:
            c.add("ETA-12", "ETA",
                  "TFT : tresorerie de cloture (%s) differente des comptes de tresorerie "
                  "de la balance N (%s)" % (_fmt(FIN), _fmt(tres_n)),
                  BLOQUANT, None, FIN - tres_n)
        if balN1 and abs(A - tres_n1) > tol:
            c.add("ETA-13", "ETA",
                  "TFT : tresorerie d'ouverture (%s) differente de la cloture des comptes "
                  "de tresorerie en N-1 (%s)" % (_fmt(A), _fmt(tres_n1)),
                  BLOQUANT, None, A - tres_n1)
        tres_bilan = v(bilan, "H37") - v(bilan, "M37")
        if tres_bilan and abs(FIN - tres_bilan) > tol:
            c.add("ETA-14", "ETA",
                  "TFT : tresorerie de cloture (%s) differente de la tresorerie nette "
                  "du bilan, actif moins passif (%s)" % (_fmt(FIN), _fmt(tres_bilan)),
                  BLOQUANT, None, FIN - tres_bilan)

    # ── Notes : rattachement du total de chaque note a la balance ───────────
    #
    # Chaque note doit se recouper avec les comptes qui l'alimentent. Un ecart
    # signale soit une saisie manuelle incoherente, soit un compte que la note
    # ne capte pas. Le cas le plus frequent est la note vide alors que les
    # comptes correspondants sont mouvementes.
    NOTES = [
        ("ETA-20", "NOTE 3A", "J30", 21000, 28000, False,
         "Immobilisations brutes"),
        ("ETA-21", "NOTE 3C", "N24", 28000, 29000, True,
         "Amortissements cumules"),
        # La NOTE 3C BIS ne couvre que les depreciations des immobilisations
        # incorporelles et corporelles (291 a 295). Celles des immobilisations
        # financieres (296 titres de participation, 297 autres) relevent de la
        # NOTE 4, d'ou la borne haute a 29600 et le controle ETA-28 plus bas.
        ("ETA-22", "NOTE 3C BIS", "N24", 29100, 29600, True,
         "Depreciations des immobilisations incorporelles et corporelles"),
        ("ETA-27", "NOTE 4", "F18", 26000, 28000, False,
         "Immobilisations financieres brutes"),
        ("ETA-28", "NOTE 4", "F19", 29600, 29800, True,
         "Depreciations des immobilisations financieres"),
        ("ETA-23", "NOTE 6", "E17", 31000, 39000, False,
         "Stocks bruts"),
        ("ETA-25", "NOTE 16A", "K19", 16000, 17000, True,
         "Emprunts et dettes financieres"),
        ("ETA-26", "NOTE 16A", "K25", 17000, 18000, True,
         "Dettes de location-acquisition"),
    ]
    for code, feuille, coord, lo, hi, absolu, libelle in NOTES:
        note = (notes or {}).get(feuille)
        if note is None:
            continue
        montant = _n(note.get(coord))
        attendu = _somme(balN, lo, hi)
        if absolu:
            attendu = abs(attendu)
        if abs(montant - attendu) <= tol:
            continue
        if abs(montant) < 0.5 and abs(attendu) >= 0.5:
            c.add(code, "ETA",
                  "%s : la note est vide alors que les comptes correspondants de la "
                  "balance totalisent %s" % (feuille, _fmt(attendu)),
                  MAJEUR, None, attendu)
        else:
            c.add(code, "ETA",
                  "%s (%s) : total de %s contre %s dans la balance"
                  % (feuille, libelle, _fmt(montant), _fmt(attendu)),
                  MAJEUR, None, montant - attendu)

    # NOTE 7 : le total brut inclut les clients crediteurs, qu'il faut retrancher
    n7 = (notes or {}).get("NOTE 7")
    if n7 is not None:
        brut = _n(n7.get("E17")) - _n(n7.get("E23"))
        attendu = _somme(balN, 41000, 42000)
        if abs(brut) >= 0.5 or abs(attendu) >= 0.5:
            if abs(brut - attendu) > tol:
                c.add("ETA-24", "ETA",
                      "NOTE 7 (Clients) : total net de %s contre %s dans la balance"
                      % (_fmt(brut), _fmt(attendu)), MAJEUR, None, brut - attendu)


# ---------------------------------------------------------------------------
# CPL  completude avant depot
# ---------------------------------------------------------------------------
def controles_cpl(c, client, exo, manual, balN):
    if not (client and str(client["ncc"] or "").strip()):
        c.add("CPL-01", "CPL",
              "NCC du client non renseigne : le fichier XML sera rejete par la DGI",
              BLOQUANT)
    d1 = exo["date_debut"] if exo else None
    d2 = exo["date_fin"] if exo else None
    if not (d1 and d2):
        c.add("CPL-02", "CPL",
              "Dates de debut et de fin d'exercice non renseignees", BLOQUANT)
    else:
        try:
            from datetime import datetime
            a = datetime.strptime(str(d1)[:10], "%Y-%m-%d")
            b = datetime.strptime(str(d2)[:10], "%Y-%m-%d")
            mois = (b.year - a.year) * 12 + (b.month - a.month) + 1
            if mois != 12:
                c.add("CPL-02", "CPL",
                      "Duree de l'exercice de %d mois : a justifier s'il ne s'agit pas "
                      "du premier exercice" % mois, JUSTIFIER)
        except Exception:
            pass
    if not balN:
        c.add("CPL-10", "CPL", "Aucune balance importee pour l'exercice N", BLOQUANT)

    r1 = (manual or {}).get("FICHE R1") or {}
    if not any(str(x).strip() for x in r1.values()):
        c.add("CPL-03", "CPL",
              "Fiche R1 non renseignee : signataire et date de signature obligatoires",
              MAJEUR)
    r2 = (manual or {}).get("FICHE R2") or {}
    if not any(str(x).strip() for x in r2.values()):
        c.add("CPL-04", "CPL",
              "Fiche R2 non renseignee : activites et code CIAP obligatoires", MAJEUR)
    r3 = (manual or {}).get("FICHE R3") or {}
    if not any(str(x).strip() for x in r3.values()):
        c.add("CPL-05", "CPL", "Fiche R3 : aucun dirigeant renseigne", MAJEUR)


# ---------------------------------------------------------------------------
# point d'entree
# ---------------------------------------------------------------------------
def run_controls(balN, balN1, bilan=None, resultat=None, tft=None, manual=None,
                 client=None, exo=None, sheets_raw=None, params=None, overrides=None,
                 mappings=None, notes=None, condenser=True):
    """Execute l'ensemble des familles et retourne la liste des anomalies.

    condenser=False renvoie le detail integral, sans regroupement : c'est ce que
    l'export CSV utilise, pour que le collaborateur dispose de la liste complete
    compte par compte meme quand la page en resume une partie."""
    p = dict(DEFAUTS)
    p.update(params or {})
    c = _Collecteur(p)

    if SENS_PLAN_ERREUR:
        c.add("MAP-11", "MAP",
              "Referentiel des sens comptables introuvable : %s. "
              "Les controles de sens (famille SEN) et la detection des comptes "
              "hors plan (INT-06) sont desactives. Verifiez que revision_sens.json "
              "accompagne bien l'application." % SENS_PLAN_ERREUR, BLOQUANT)

    try:
        controles_int(c, balN, balN1, overrides, mappings)
    except Exception as e:
        c.add("INT-ERR", "INT", "Erreur interne du controle : %s" % e, INFO)
    for fn, args, fam in (
        (controles_sen, (balN, overrides, mappings), "SEN"),
        (controles_att, (balN,), "ATT"),
        (controles_fis, (balN, resultat), "FIS"),
        (controles_soc, (balN,), "SOC"),
        (controles_imm, (balN, balN1, manual), "IMM"),
        (controles_stk, (balN, balN1), "STK"),
        (controles_var, (balN, balN1, bilan), "VAR"),
        (controles_eta, (bilan, resultat, tft, balN, balN1, notes), "ETA"),
        (controles_cpl, (client, exo, manual, balN), "CPL"),
    ):
        try:
            fn(c, *args)
        except Exception as e:
            c.add("%s-ERR" % fam, fam, "Erreur interne du controle : %s" % e, INFO)
    if sheets_raw:
        try:
            controles_map(c, sheets_raw, balN, balN1)
        except Exception as e:
            c.add("MAP-ERR", "MAP", "Erreur interne du controle : %s" % e, INFO)

    items = _condenser(c.items) if condenser else list(c.items)
    items.sort(key=lambda a: (SEVERITE_ORDRE.get(a["severite"], 9),
                              a["famille"], a["code"], a["compte"] or ""))
    return items


def _condenser(items):
    """Anti-bruit : au-dela d'un plafond, ne garder que les plus gros montants
    et resumer le reste en une ligne. Un module qui remonte 400 anomalies ne
    sera pas utilise ; ce filtre est ce qui rend la page exploitable."""
    groupes = {}
    for a in items:
        groupes.setdefault(a["code"], []).append(a)
    out = []
    for code, lot in groupes.items():
        # Un plafond bas ne doit jamais supprimer TOUT le detail : on montre
        # toujours au moins CONDENSATION_MINI occurrences, les plus grosses,
        # avant de resumer le reste. Sans ce plancher, un code plafonne a zero
        # n'affichait qu'une ligne de synthese sans aucun compte.
        plafond = max(PLAFOND.get(code, PLAFOND_DEFAUT), CONDENSATION_MINI)
        if len(lot) <= plafond:
            out.extend(lot)
            continue
        lot.sort(key=lambda a: -abs(_n(a.get("montant"))))
        garde = lot[:plafond]
        reste = lot[plafond:]
        out.extend(garde)
        total = sum(_n(a.get("montant")) for a in reste)
        modele = lot[0]
        out.append({
            "code": code,
            "famille": modele["famille"],
            "libelle": "%d autre%s compte%s concerne%s par ce controle%s "
                       "- utilisez l'export CSV pour la liste complete"
                       % (len(reste),
                          "s" if len(reste) > 1 else "",
                          "s" if len(reste) > 1 else "",
                          "s" if len(reste) > 1 else "",
                          " (total %s)" % _fmt(total) if total else ""),
            "severite": INFO,
            "compte": None,
            "montant": round(total, 2) if total else None,
            "agrege": len(reste),
        })
    return out


def synthese(anomalies):
    """Compteurs par severite, pour l'entete de la page et le tableau de bord."""
    out = {BLOQUANT: 0, MAJEUR: 0, JUSTIFIER: 0, INFO: 0}
    for a in anomalies:
        out[a["severite"]] = out.get(a["severite"], 0) + 1
    out["total"] = len(anomalies)
    return out
