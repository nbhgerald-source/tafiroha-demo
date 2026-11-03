"""
Génération du fichier XML EDI (e-impôts DGI) à partir des données TAFIROHA.

Structure XML attendue (définie dans le schéma EDI_Mappage du fichier Excel) :

    <EDI>
      <informations>
        <type>NO</type>
        <ncc>...</ncc>
        <exercice>...</exercice>
      </informations>
      <champsTableauxFixes>
        <champTableauFixe>
          <code>NO_FR1_ZA1_1</code>
          <valeur>'12345</valeur>
        </champTableauFixe>
        …
      </champsTableauxFixes>
      <champsTableauxVariables>
        <champTableauVariable>
          <colonne>NO_NOTE8A2_1</colonne>
          <ligne>1</ligne>
          <valeur>42000</valeur>
        </champTableauVariable>
        …
      </champsTableauxVariables>
    </EDI>

Logique de formatage des valeurs (équivalent GetValeur VBA) :
  - Texte / code alphanumérique  → "'texte"   (apostrophe préfixe)
  - Code numérique "0xxx"        → "'0xxx"    (apostrophe préfixe)
  - Entier                       → "1234"
  - Décimal                      → "1234.5678" (4 décimales max)

Seules les cellules non vides sont exportées (conformément au comportement VBA).
"""

import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta

# ── Chargement du mapping XML (une fois au démarrage du module) ──────────────

if getattr(sys, "frozen", False):
    _BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(_BASE, "xml_mapping.json"), encoding="utf-8") as _fh:
    _MAP = json.load(_fh)

# tableaux : dict  code_tableau  → {onglet, premiere, premiere_col, premiere_row, derniere, …}
# champs_fixes   : list of {code_tableau, code_champ, pos_ligne, pos_col, onglet, premiere_col, premiere_row}
# champs_variables: list of {code_tableau, code_colonne, pos_col, onglet, premiere_col, premiere_row, derniere}
_TABLEAUX       = _MAP["tableaux"]
_CHAMPS_FIXES   = _MAP["champs_fixes"]
_CHAMPS_VARS    = _MAP["champs_variables"]

# Feuilles calculées par calc_engine (clés de sheets_raw.json)
_CALC_SHEETS = {
    "BILAN", "RESULTAT", "TFT",
    "NOTE 3A", "NOTE 3B", "NOTE 3C", "NOTE 3C BIS", "NOTE 3D", "NOTE 3E",
    "NOTE 4", "NOTE 5", "NOTE 6", "NOTE 7", "NOTE 8", "NOTE 9",
    "NOTE 10", "NOTE 11", "NOTE 14", "NOTE 15A", "NOTE 16A",
    "NOTE 17", "NOTE 18", "NOTE 19", "NOTE 20", "NOTE 21",
    "NOTE 22", "NOTE 23", "NOTE 24", "NOTE 25", "NOTE 26",
    "NOTE 27A", "NOTE 28", "NOTE 29",
}

# ── Conversions colonne ──────────────────────────────────────────────────────

def _num_to_col(n: int) -> str:
    """Numéro de colonne (1-based) → lettres Excel : 1→A, 26→Z, 27→AA …"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _derniere_row(s: str) -> int:
    """Extrait le numéro de ligne depuis une adresse cellule style 'AH29'."""
    m = re.match(r"[A-Z]+(\d+)", s)
    return int(m.group(1)) if m else 0


# ── Passerelle FICHE R1 / R2 / R3 ───────────────────────────────────────────
#
# Les feuilles d'identification ne sont pas des feuilles de calcul : elles sont
# saisies dans l'application, qui les range dans note_texte sous des NOMS de
# champ (sheet="SOMMAIRE", champ="exercice_debut") et non sous des coordonnees
# Excel. _get_value, lui, interroge note_texte[onglet]["N8"] : les deux
# adressages ne se rencontrent jamais et tout le bloc FR1/FR2/FR3 sortait vide
# du fichier EDI — la plateforme e-impots recevait une liasse sans en-tete.
#
# Cette table retablit la correspondance, code DGI par code DGI. Les cellules
# et les libelles proviennent de ListeChampsTableauxFixes /
# ListeColonnesTableauxVariables du classeur officiel.

def _date_fr(v):
    """Normalise une date en jj/mm/aaaa (format attendu par GetValeur VBA)."""
    s = str(v or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


def _codes_fiche_r2(somm):
    """ZX / ZY / ZZ1 : codes statistiques recalcules depuis les choix faits sur
    SOMMAIRE, comme les formules INDEX/MATCH du classeur (FICHE R2 Q9/Q11/Q13).
    Les tables de choix vivent dans app.py ; import tardif pour ne pas creer de
    cycle (app importe deja xml_engine)."""
    try:
        import app as _app
    except Exception:
        return "", "", ""
    zx = ""
    forme = somm.get("forme_juridique", "")
    choix = getattr(_app, "FORME_JURIDIQUE_CHOICES", [])
    if forme in choix:
        agrement = "1" if somm.get("agrement_prioritaire") == "Oui" else "0"
        zx = agrement + str(choix.index(forme))
    zy = ""
    regime = somm.get("regime_fiscal", "")
    choix = getattr(_app, "REGIME_FISCAL_CHOICES", [])
    if regime in choix:
        zy = str(choix.index(regime) + 1)
    zz1 = getattr(_app, "PAYS_SIEGE_CODES", {}).get(somm.get("pays_siege", ""), "")
    return zx, zy, zz1


def _fiches_identification(client: dict, exo: dict, ntexte: dict, note3: dict):
    """Retourne (fixes, variables) :

        fixes     : {code_champ: (valeur, est_texte)}
        variables : {(code_colonne, no_ligne): (valeur, est_texte)}

    Seules les valeurs reellement saisies sont renvoyees : une entree vide est
    omise, exactement comme le fait la macro officielle qui ne retient que les
    cellules dont le texte affiche n'est pas vide.
    """
    somm = ntexte.get("SOMMAIRE", {}) or {}
    r1 = ntexte.get("FICHE R1", {}) or {}
    r2 = ntexte.get("FICHE R2", {}) or {}
    r3 = ntexte.get("FICHE R3", {}) or {}

    T, N = True, False          # est_texte
    fixes = {}

    def put(code, valeur, est_texte=T):
        s = "" if valeur is None else str(valeur).strip()
        if s:
            fixes[code] = (s, est_texte)

    # ── FICHE R1 ────────────────────────────────────────────────────────────
    put("NO_FR1_ZA1_1", _date_fr(somm.get("exercice_debut") or exo.get("date_debut")))
    put("NO_FR1_ZA2_1", _date_fr(somm.get("exercice_fin") or exo.get("date_fin")))
    put("NO_FR1_ZB_1",  _date_fr(somm.get("date_arret_comptes")))
    # ZC "exercice precedent clos le" : l'application ne saisit pas cette date.
    # Elle se deduit sans ambiguite du debut de l'exercice courant — la veille.
    # ZD (duree de l'exercice precedent en mois) n'a aucune source : laisse vide
    # plutot que devine.
    debut = _date_fr(somm.get("exercice_debut") or exo.get("date_debut"))
    try:
        veille = datetime.strptime(debut, "%d/%m/%Y") - timedelta(days=1)
        put("NO_FR1_ZC_1", veille.strftime("%d/%m/%Y"))
    except ValueError:
        pass
    put("NO_FR1_ZE1_1", somm.get("greffe"))
    put("NO_FR1_ZE2_1", somm.get("rccm"))
    put("NO_FR1_ZF_1",  somm.get("repertoire_entreprises"))
    put("NO_FR1_ZG_1",  somm.get("caisse_sociale"))
    put("NO_FR1_ZH_1",  somm.get("code_importateur"))
    put("NO_FR1_ZI_1",  somm.get("code_activite"))
    put("NO_FR1_ZK1_1", somm.get("email"))
    put("NO_FR1_ZK2_1", somm.get("boite_postale_code"))
    put("NO_FR1_ZK4_1", somm.get("boite_postale_ville"))
    put("NO_FR1_ZL_1",  client.get("adresse"))
    # L'application saisit la personne a contacter et l'expert-comptable en une
    # seule zone de texte, la DGI les attend eclates (nom / adresse / tel /
    # mail / qualite). On porte la saisie dans la zone "Nom" plutot que de la
    # perdre ; les autres zones restent vides tant qu'elles ne sont pas saisies
    # separement dans l'application.
    put("NO_FR1_ZO1_1", r1.get("contact_nom"))
    put("NO_FR1_ZP1_1", r1.get("expert_comptable"))
    put("NO_FR1_ZT_1",  r1.get("signataire_nom"))
    put("NO_FR1_ZU_1",  r1.get("signataire_qualite"))
    put("NO_FR1_ZV_1",  _date_fr(r1.get("signataire_date")))

    # ── FICHE R2 ────────────────────────────────────────────────────────────
    zx, zy, zz1 = _codes_fiche_r2(somm)
    put("NO_FR2A_ZX_1",  zx)
    put("NO_FR2A_ZY_1",  zy)
    put("NO_FR2A_ZZ1_1", zz1)
    put("NO_FR2A_ZZ2_1", r2.get("nb_etabl_pays"), N)
    put("NO_FR2A_ZZ3_1", r2.get("nb_etabl_hors_pays"), N)
    put("NO_FR2A_ZZ4_1", somm.get("annee_premier_exercice"), N)
    put("NO_FR2A_1_1",   r2.get("type_controle"))
    put("NO_FR2C_1_1",   r2.get("type_montant_activite"))

    # ── FICHE R2B — activites (tableau variable) ────────────────────────────
    variables = {}
    montants, total = {}, 0.0
    nb_act = int(getattr(__import__("sys").modules.get("app"), "FICHE_R2_ACTIVITES_ROWS", 6) or 6)
    for i in range(1, nb_act + 1):
        try:
            m = float(str(r2.get("activite_%d_montant" % i, "") or 0).replace(" ", "").replace(",", "."))
        except ValueError:
            m = 0.0
        montants[i] = m
        total += m
    ligne = 0
    for i in range(1, nb_act + 1):
        designation = (r2.get("activite_%d_designation" % i, "") or "").strip()
        code_act = (r2.get("activite_%d_code" % i, "") or "").strip()
        if not (designation or code_act or montants[i]):
            continue
        ligne += 1
        if designation:
            variables[("NO_FR2B_1", ligne)] = (designation, T)
        if code_act:
            variables[("NO_FR2B_2", ligne)] = (code_act, T)
        if montants[i]:
            variables[("NO_FR2B_3", ligne)] = (montants[i], N)
            if total:
                variables[("NO_FR2B_4", ligne)] = (round(montants[i] / total, 4), N)
    if total:
        fixes["NO_FR2D_1_1"] = (total, N)
        fixes["NO_FR2D_1_2"] = (1, N)

    # ── FICHE R3 — dirigeants et conseil d'administration ───────────────────
    def _nb(cle, defaut):
        try:
            return int(float(note3.get("FICHE R3", {}).get(cle, defaut) or defaut))
        except (TypeError, ValueError):
            return defaut

    ligne = 0
    for i in range(1, _nb("_N_DIRIGEANT", 5) + 1):
        champs = {
            "NO_FR3A_1": r3.get("dirigeant_%d_nom" % i),
            "NO_FR3A_2": r3.get("dirigeant_%d_nationalite" % i),
            "NO_FR3A_4": r3.get("dirigeant_%d_qualite" % i),
            "NO_FR3A_5": r3.get("dirigeant_%d_nif" % i),
            "NO_FR3A_6": r3.get("dirigeant_%d_adresse" % i),
        }
        if not any((v or "").strip() for v in champs.values()):
            continue
        ligne += 1
        for code, v in champs.items():
            if (v or "").strip():
                variables[(code, ligne)] = (str(v).strip(), T)

    ligne = 0
    for i in range(1, _nb("_N_ADMIN", 5) + 1):
        champs = {
            "NO_FR3B_1": r3.get("admin_%d_nom" % i),
            "NO_FR3B_2": r3.get("admin_%d_structure" % i),
            "NO_FR3B_3": r3.get("admin_%d_qualite" % i),
            "NO_FR3B_4": r3.get("admin_%d_adresse" % i),
        }
        if not any((v or "").strip() for v in champs.values()):
            continue
        ligne += 1
        for code, v in champs.items():
            if (v or "").strip():
                variables[(code, ligne)] = (str(v).strip(), T)

    return fixes, variables


# ── Lecture de la valeur d'une cellule ──────────────────────────────────────

def _get_value(computed: dict, note3: dict, ntexte: dict, onglet: str, row: int, col: int):
    """
    Retourne (valeur, est_texte).

    - ACTIF  → coordonnée identique dans BILAN (col non décalée)
    - PASSIF → colonne +5 dans BILAN
    - Feuilles calculées → computed[onglet][coord]
    - Feuilles manuelles → note_texte d'abord (texte), puis note3_manuel (numérique)
      Si la cellule n'a jamais été saisie → (None, False) → sera ignorée
    """
    coord = _num_to_col(col) + str(row)

    if onglet == "ACTIF":
        v = computed.get("BILAN", {}).get(coord, 0)
        return v, False

    if onglet == "PASSIF":
        bilan_coord = _num_to_col(col + 5) + str(row)
        v = computed.get("BILAN", {}).get(bilan_coord, 0)
        return v, False

    if onglet in computed:
        # Feuille calculée — on inclut même les 0 (cellule avec formule)
        v = computed[onglet].get(coord, 0)
        return v, False

    # Feuille manuelle
    if ntexte.get(onglet, {}).get(coord) not in (None, ""):
        return ntexte[onglet][coord], True

    if note3.get(onglet, {}).get(coord) is not None:
        return note3[onglet][coord], False

    return None, False  # cellule vide → à ignorer


# ── Formatage de la valeur (équivalent GetValeur VBA) ───────────────────────

def _format_value(v, is_text: bool):
    """
    Retourne la chaîne à écrire dans <valeur>, ou None si la cellule est vide.
    """
    if v is None:
        return None

    if is_text:
        s = str(v).strip()
        return ("'" + s) if s else None

    if isinstance(v, (int, float)):
        if isinstance(v, bool):
            return None
        f = float(v)
        if f == 0.0:
            return "0"
        if f == int(f):
            return str(int(f))
        return str(round(f, 4))

    # Valeur stockée comme chaîne (ne devrait pas arriver pour note3)
    s = str(v).strip()
    if not s:
        return None
    try:
        n = float(s)
        if s.startswith("0") and len(s) > 1:   # Code "01234" → texte
            return "'" + s
        return str(int(n)) if n == int(n) else str(round(n, 4))
    except ValueError:
        return ("'" + s) if s else None


# ── Échappement XML ──────────────────────────────────────────────────────────

def _esc(s) -> str:
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


# ── Génération XML principale ────────────────────────────────────────────────

def generate_xml(conn, exercice_id: int, client: dict, exo: dict, periode: str = "N") -> bytes:
    """
    Génère le XML EDI (bytes UTF-8) pour l'exercice donné.

    Paramètres :
        conn         : connexion SQLite (sqlite3.Connection)
        exercice_id  : identifiant de l'exercice
        client       : dict avec au moins 'ncc' et 'raison_sociale'
        exo          : dict avec au moins 'annee'
        periode      : période comptable principale ("N" par défaut)

    Retourne les octets du fichier XML prêt à être envoyé/enregistré.
    """
    import calc_engine as ce

    # 1. Chargement des données ──────────────────────────────────────────────

    # Codification du client : le fichier XML doit porter les memes montants
    # que les etats affiches a l'ecran, donc appliquer le meme rattachement.
    def _mappings():
        try:
            r = conn.execute("SELECT client_id FROM exercices WHERE id=?",
                             (exercice_id,)).fetchone()
            if not r:
                return {}
            return {int(x["cle_source"]): int(x["cle_cible"]) for x in conn.execute(
                "SELECT cle_source, cle_cible FROM client_compte_mapping WHERE client_id=?",
                (r["client_id"],)).fetchall()}
        except Exception:
            return {}

    _map = _mappings()

    def _bal(per):
        rows = conn.execute(
            "SELECT * FROM balance_lignes WHERE exercice_id=? AND periode=? ORDER BY compte",
            (exercice_id, per),
        ).fetchall()
        return ce.build_balance_rows([dict(r) for r in rows], _map)

    def _tft(per):
        rows = conn.execute(
            "SELECT * FROM tft_detail_lignes WHERE exercice_id=? AND periode=? ORDER BY compte",
            (exercice_id, per),
        ).fetchall()
        return ce.build_balance_rows([dict(r) for r in rows], _map)

    def _has_per(per):
        return bool(conn.execute(
            "SELECT 1 FROM balance_lignes WHERE exercice_id=? AND periode=? LIMIT 1",
            (exercice_id, per),
        ).fetchone())

    periode_n1 = "N-1" if _has_per("N-1") else ("N1" if _has_per("N1") else None)

    balN   = _bal(periode)
    balN1  = _bal(periode_n1) if periode_n1 else []
    tftn   = _tft(periode)
    tftn1  = _tft(periode_n1) if periode_n1 else []

    # note3_manuel : {sheet: {coord: valeur_réel}}
    note3_rows = conn.execute(
        "SELECT sheet, coord, valeur FROM note3_manuel WHERE exercice_id=?", (exercice_id,)
    ).fetchall()
    note3: dict = {}
    for r in note3_rows:
        note3.setdefault(r["sheet"], {})[r["coord"]] = r["valeur"]

    # note_texte : {sheet: {champ: texte}}
    ntexte_rows = conn.execute(
        "SELECT sheet, champ, texte FROM note_texte WHERE exercice_id=?", (exercice_id,)
    ).fetchall()
    ntexte: dict = {}
    for r in ntexte_rows:
        ntexte.setdefault(r["sheet"], {})[r["champ"]] = r["texte"]

    # 2. Calcul des feuilles ─────────────────────────────────────────────────

    # BILAN, RESULTAT, TFT : via compute_sheet (TFT nécessite les détails TFT)
    bilan_cache    = ce.compute_sheet("BILAN",    balN, balN1) if balN else {}
    resultat_cache = ce.compute_sheet("RESULTAT", balN, balN1) if balN else {}
    tft_cache      = ce.compute_sheet("TFT",      balN, balN1, tftn, tftn1) if balN else {}

    # Autres feuilles calculées : via compute_workbook (gère les références croisées)
    other_calc = [s for s in ce.SHEETS_RAW.keys() if s not in ("BILAN", "RESULTAT", "TFT")]
    manual_for_engine = {sh: note3[sh] for sh in other_calc if sh in note3}
    notes_cache = ce.compute_workbook(
        other_calc, balN, balN1,
        manual=manual_for_engine,
        row_range=(1, 200),
    ) if balN else {}

    # Vue consolidée : computed[onglet][coord] = valeur
    computed: dict = dict(notes_cache)
    computed["BILAN"]    = bilan_cache
    computed["RESULTAT"] = resultat_cache
    computed["TFT"]      = tft_cache

    # 3. Construction du XML ─────────────────────────────────────────────────

    ncc    = (client.get("ncc") or "").strip()
    annee  = str(exo.get("annee", ""))

    out = ['<?xml version="1.0" encoding="UTF-8"?>']
    out.append("<EDI>")

    # --- informations ---
    out.append("  <informations>")
    out.append("    <type>NO</type>")
    out.append(f"    <ncc>{_esc(ncc)}</ncc>")
    out.append(f"    <exercice>{_esc(annee)}</exercice>")
    out.append("  </informations>")

    # Identification (FICHE R1 / R2 / R3) : ces feuilles sont saisies, pas
    # calculees, et l'application les range par nom de champ et non par
    # coordonnee Excel — elles ne peuvent donc pas passer par _get_value.
    fiches_fixes, fiches_vars = _fiches_identification(client, exo, ntexte, note3)

    # --- champsTableauxFixes ---
    out.append("  <champsTableauxFixes>")
    for c in _CHAMPS_FIXES:
        if c["code_champ"] in fiches_fixes:
            v, is_text = fiches_fixes[c["code_champ"]]
        else:
            row = c["premiere_row"] + c["pos_ligne"] - 1
            col = c["premiere_col"] + c["pos_col"] - 1
            v, is_text = _get_value(computed, note3, ntexte, c["onglet"], row, col)
        fv = _format_value(v, is_text)
        if fv is not None:
            out.append("    <champTableauFixe>")
            out.append(f"      <code>{_esc(c['code_champ'])}</code>")
            out.append(f"      <valeur>{_esc(fv)}</valeur>")
            out.append("    </champTableauFixe>")
    out.append("  </champsTableauxFixes>")

    # --- champsTableauxVariables ---
    out.append("  <champsTableauxVariables>")

    # Regrouper les colonnes par tableau pour itérer les lignes une seule fois
    cols_by_tab: dict = defaultdict(list)
    for c in _CHAMPS_VARS:
        cols_by_tab[c["code_tableau"]].append(c)

    for code_tab, cols in cols_by_tab.items():
        tab_info    = _TABLEAUX.get(code_tab, {})
        premiere_row = tab_info.get("premiere_row", cols[0]["premiere_row"])
        derniere_str = tab_info.get("derniere", cols[0].get("derniere", ""))
        last_row     = _derniere_row(derniere_str) if derniere_str else premiere_row
        n_rows       = last_row - premiere_row + 1

        # Les tableaux d'identification ont un nombre de lignes propre : il
        # vient de la saisie, pas de l'etendue d'une plage du classeur.
        lignes_fiche = sorted({l for (code, l) in fiches_vars
                               if code in {c["code_colonne"] for c in cols}})
        plage = lignes_fiche if lignes_fiche else range(1, n_rows + 1)

        for i_ligne in plage:
            row = premiere_row + i_ligne - 1
            for c in cols:
                cle = (c["code_colonne"], i_ligne)
                if cle in fiches_vars:
                    v, is_text = fiches_vars[cle]
                elif lignes_fiche:
                    continue          # tableau d'identification : rien de plus
                else:
                    col = c["premiere_col"] + c["pos_col"] - 1
                    v, is_text = _get_value(computed, note3, ntexte, c["onglet"], row, col)
                fv = _format_value(v, is_text)
                if fv is not None:
                    out.append("    <champTableauVariable>")
                    out.append(f"      <colonne>{_esc(c['code_colonne'])}</colonne>")
                    out.append(f"      <ligne>{i_ligne}</ligne>")
                    out.append(f"      <valeur>{_esc(fv)}</valeur>")
                    out.append("    </champTableauVariable>")

    out.append("  </champsTableauxVariables>")
    out.append("</EDI>")

    return "\n".join(out).encode("utf-8")


def xml_filename(client: dict, exo: dict) -> str:
    """Retourne le nom de fichier : NO-{annee}-{ncc}-{ddmmyyyy-hhmmss}.xml"""
    ncc   = re.sub(r"[^A-Za-z0-9]", "", (client.get("ncc") or "INCONNU"))
    annee = str(exo.get("annee", ""))
    ts    = datetime.now().strftime("%d%m%Y-%H%M%S")
    return f"NO-{annee}-{ncc}-{ts}.xml"
