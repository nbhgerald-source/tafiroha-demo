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


# ── Notes litteraires : meme defaut que les fiches d'identification ─────────
#
# NOTE 2, 35, 38, 39 et la feuille COMMENTAIRE sont saisies en texte libre et
# rangees dans note_texte sous des noms de champ. Elles ne pouvaient donc pas
# passer par _get_value, qui cherche une coordonnee Excel. Correspondance
# etablie code par code sur les libelles de ListeChampsTableauxFixes ; pour
# COMMENT, le referentiel donne directement le tableau vise en libelle de
# ligne (NO_NOTE3A, NO_SUPPL6, ...), la table ci-dessous le transpose vers la
# feuille et la zone correspondantes de l'application.

_NOTES_TEXTE_MAP = {
    # NOTE 2 — informations obligatoires
    "NO_NOTE2_1_1": ("NOTE 2", "conformite"),
    "NO_NOTE2_2_1": ("NOTE 2", "regles_methodes"),
    "NO_NOTE2_3_1": ("NOTE 2", "derogations"),
    "NO_NOTE2_4_1": ("NOTE 2", "infos_complementaires"),
    # NOTE 35 — informations sociales, environnementales et societales
    "NO_NOTE35_2_1":  ("NOTE 35", "emploi"),
    "NO_NOTE35_3_1":  ("NOTE 35", "relations_sociales"),
    "NO_NOTE35_4_1":  ("NOTE 35", "sante_securite"),
    "NO_NOTE35_5_1":  ("NOTE 35", "formation"),
    "NO_NOTE35_6_1":  ("NOTE 35", "egalite"),
    "NO_NOTE35_8_1":  ("NOTE 35", "env_general"),
    "NO_NOTE35_9_1":  ("NOTE 35", "pollution_dechets"),
    "NO_NOTE35_10_1": ("NOTE 35", "ressources"),
    "NO_NOTE35_11_1": ("NOTE 35", "climat"),
    "NO_NOTE35_12_1": ("NOTE 35", "biodiversite"),
    "NO_NOTE35_14_1": ("NOTE 35", "impact_territorial"),
    "NO_NOTE35_15_1": ("NOTE 35", "relations_parties"),
    "NO_NOTE35_16_1": ("NOTE 35", "sous_traitance"),
    # NOTE 38 — evenements posterieurs a la cloture
    "NO_NOTE38A_1_1": ("NOTE 38", "date_arrete"),
    "NO_NOTE38A_2_1": ("NOTE 38", "organe_publication"),
    "NO_NOTE38B_1_1": ("NOTE 38", "avec_ajustement"),
    "NO_NOTE38B_2_1": ("NOTE 38", "sans_ajustement"),
    "NO_NOTE38B_3_1": ("NOTE 38", "continuite"),
    # NOTE 39 — changements de methodes, d'estimations, corrections d'erreurs
    "NO_NOTE39_1_1": ("NOTE 39", "methodes"),
    "NO_NOTE39_2_1": ("NOTE 39", "estimations"),
    "NO_NOTE39_3_1": ("NOTE 39", "corrections"),
    # COMMENTAIRE — une zone par note
    "NO_COMMENT_11_1": ("NOTE 1", "commentaire_0"),
    "NO_COMMENT_13_1": ("NOTE 3A", "commentaire_0"),
    "NO_COMMENT_14_1": ("NOTE 3B", "commentaire_0"),
    "NO_COMMENT_15_1": ("NOTE 3C", "commentaire_0"),
    "NO_COMMENT_16_1": ("NOTE 3C BIS", "commentaire_0"),
    "NO_COMMENT_17_1": ("NOTE 3D", "commentaire_0"),
    "NO_COMMENT_22_1": ("NOTE 4", "commentaire_0"),
    "NO_COMMENT_23_1": ("NOTE 5", "commentaire_0"),
    "NO_COMMENT_24_1": ("NOTE 5", "commentaire_1"),
    "NO_COMMENT_25_1": ("NOTE 6", "commentaire_0"),
    "NO_COMMENT_26_1": ("NOTE 7", "commentaire_0"),
    "NO_COMMENT_27_1": ("NOTE 8", "commentaire_0"),
    "NO_COMMENT_34_1": ("NOTE 9", "commentaire_0"),
    "NO_COMMENT_35_1": ("NOTE 10", "commentaire_0"),
    "NO_COMMENT_36_1": ("NOTE 11", "commentaire_0"),
    "NO_COMMENT_38_1": ("NOTE 12", "commentaire_0"),
    "NO_COMMENT_40_1": ("NOTE 12", "commentaire_1"),
    "NO_COMMENT_43_1": ("NOTE 13", "commentaire_0"),
    "NO_COMMENT_44_1": ("NOTE 14", "commentaire_0"),
    "NO_COMMENT_45_1": ("NOTE 15A", "commentaire_0"),
    "NO_COMMENT_46_1": ("NOTE 15B", "commentaire_0"),
    "NO_COMMENT_47_1": ("NOTE 16A", "commentaire_0"),
    "NO_COMMENT_48_1": ("NOTE 16B", "commentaire_0"),
    "NO_COMMENT_49_1": ("NOTE 16B", "commentaire_1"),
    "NO_COMMENT_50_1": ("NOTE 16B", "commentaire_2"),
    "NO_COMMENT_51_1": ("NOTE 16B BIS", "commentaire_0"),
    "NO_COMMENT_52_1": ("NOTE 16B BIS", "commentaire_1"),
    "NO_COMMENT_55_1": ("NOTE 16C", "commentaire_0"),
    "NO_COMMENT_56_1": ("NOTE 17", "commentaire_0"),
    "NO_COMMENT_57_1": ("NOTE 18", "commentaire_0"),
    "NO_COMMENT_58_1": ("NOTE 19", "commentaire_0"),
    "NO_COMMENT_59_1": ("NOTE 20", "commentaire_0"),
    "NO_COMMENT_61_1": ("NOTE 21", "commentaire_0"),
    "NO_COMMENT_62_1": ("NOTE 22", "commentaire_0"),
    "NO_COMMENT_63_1": ("NOTE 23", "commentaire_0"),
    "NO_COMMENT_64_1": ("NOTE 24", "commentaire_0"),
    "NO_COMMENT_65_1": ("NOTE 25", "commentaire_0"),
    "NO_COMMENT_66_1": ("NOTE 26", "commentaire_0"),
    "NO_COMMENT_67_1": ("NOTE 27A", "commentaire_0"),
    "NO_COMMENT_71_1": ("NOTE 27B", "commentaire_0"),
    "NO_COMMENT_72_1": ("NOTE 28", "commentaire_0"),
    "NO_COMMENT_73_1": ("NOTE 29", "commentaire_0"),
    "NO_COMMENT_76_1": ("NOTE 30", "commentaire_0"),
    "NO_COMMENT_106_1": ("SUPPL6", "commentaire_0"),
}

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
    notes_texte = {}
    for code, (feuille, champ) in _NOTES_TEXTE_MAP.items():
        v = (ntexte.get(feuille, {}) or {}).get(champ)
        if v is not None and str(v).strip():
            notes_texte[code] = (str(v).strip(), True)
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
    # ZC / ZD : cloture et duree de l'exercice precedent. Le classeur les tire
    # de Sommaire (Q37 et MONTH(Q37)-MONTH(M37)+1, vide quand l'exercice est le
    # premier). A defaut de saisie, ZC se deduit sans ambiguite : la veille du
    # debut de l'exercice courant.
    debut = _date_fr(somm.get("exercice_debut") or exo.get("date_debut"))
    prec_fin = _date_fr(somm.get("exercice_precedent_fin"))
    prec_debut = _date_fr(somm.get("exercice_precedent_debut"))
    if not prec_fin:
        try:
            prec_fin = (datetime.strptime(debut, "%d/%m/%Y")
                        - timedelta(days=1)).strftime("%d/%m/%Y")
        except ValueError:
            prec_fin = ""
    put("NO_FR1_ZC_1", prec_fin)
    try:
        premiere = str(somm.get("annee_premier_exercice") or "").strip()
        an_courant = datetime.strptime(debut, "%d/%m/%Y").year
        if not (premiere and int(premiere) == an_courant):
            d1 = datetime.strptime(prec_debut, "%d/%m/%Y")
            d2 = datetime.strptime(prec_fin, "%d/%m/%Y")
            put("NO_FR1_ZD_1", (d2.year - d1.year) * 12 + d2.month - d1.month + 1, N)
    except ValueError:
        pass
    put("NO_FR1_ZE1_1", somm.get("greffe"))
    put("NO_FR1_ZE2_1", somm.get("rccm"))
    put("NO_FR1_ZF_1",  somm.get("repertoire_entreprises"))
    put("NO_FR1_ZG_1",  somm.get("caisse_sociale"))
    put("NO_FR1_ZH_1",  somm.get("code_importateur"))
    put("NO_FR1_ZI_1",  somm.get("code_activite"))
    put("NO_FR1_ZK1_1", somm.get("email"))
    # La DGI separe le code postal, le numero de boite et la ville. Les
    # dossiers saisis avant l'ajout du champ "numero" portent souvent la
    # mention complete ("18 BP 2383") dans le code : on la redecoupe.
    bp_code = str(somm.get("boite_postale_code") or "").strip()
    bp_num = str(somm.get("boite_postale") or "").strip()
    if not bp_num:
        m = re.match(r"^\s*(\d+)\s*BP\s*(\S+)\s*$", bp_code, re.IGNORECASE)
        if m:
            bp_code, bp_num = m.group(1), m.group(2)
    put("NO_FR1_ZK2_1", bp_code)
    put("NO_FR1_ZK3_1", bp_num)
    put("NO_FR1_ZK4_1", somm.get("boite_postale_ville"))
    put("NO_FR1_ZL_1",  client.get("adresse"))

    # Personne a contacter (ZO1..ZO5) et expert-comptable (ZP1..ZP4). Les
    # sous-zones saisies font foi ; pour les dossiers anterieurs, qui ne
    # disposaient que d'une zone libre, on repartit ce qui est reconnaissable
    # — un e-mail contient "@", un telephone est une suite de chiffres — et le
    # premier fragment restant tient lieu de nom.
    def _eclater(nom, adresse, tel, mail, qualite=None, avec_qualite=False):
        vals = {"nom": nom, "adresse": adresse, "tel": tel,
                "mail": mail, "qualite": qualite}
        # Les sous-zones sont vides : la zone "nom" porte peut-etre encore une
        # saisie libre du type "NOM, telephone, e-mail, qualite".
        if not any((v or "").strip() for v in (adresse, tel, mail, qualite)):
            reste = [p.strip() for p in str(nom or "").split(",") if p.strip()]
            vals["nom"] = ""
            garde = []
            for p in reste:
                if "@" in p and not vals["mail"]:
                    vals["mail"] = p
                elif re.fullmatch(r"[\d\s.+()-]{8,}", p) and not vals["tel"]:
                    vals["tel"] = p
                else:
                    garde.append(p)
            if garde:
                vals["nom"] = garde.pop(0)
            if garde and avec_qualite:
                vals["qualite"] = garde.pop()
            if garde:
                vals["adresse"] = ", ".join(garde)
        return vals

    c = _eclater(r1.get("contact_nom"), r1.get("contact_adresse"),
                 r1.get("contact_telephone"), r1.get("contact_email"),
                 r1.get("contact_qualite"), avec_qualite=True)
    put("NO_FR1_ZO1_1", c["nom"]);     put("NO_FR1_ZO2_1", c["adresse"])
    put("NO_FR1_ZO3_1", c["tel"]);     put("NO_FR1_ZO4_1", c["mail"])
    put("NO_FR1_ZO5_1", c["qualite"])

    e = _eclater(r1.get("expert_comptable"), r1.get("expert_adresse"),
                 r1.get("expert_telephone"), r1.get("expert_email"))
    put("NO_FR1_ZP1_1", e["nom"]);     put("NO_FR1_ZP2_1", e["adresse"])
    put("NO_FR1_ZP3_1", e["tel"]);     put("NO_FR1_ZP4_1", e["mail"])

    put("NO_FR1_ZN_1",  r1.get("capacite_production"), N)
    put("NO_FR1_ZS_1",  r1.get("approuves_par"))
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
    # Le nombre d'activites est libre ; la ligne "Divers" est rangee hors
    # numerotation (activite_divers_*) et passe toujours en dernier, comme sur
    # la feuille d'origine.
    variables = {}

    def _montant(cle):
        try:
            return float(str(r2.get(cle, "") or 0).replace(" ", "").replace(",", "."))
        except ValueError:
            return 0.0

    indices = sorted({int(k.split("_")[1]) for k in r2
                      if k.startswith("activite_") and k.split("_")[1].isdigit()})
    lignes_act = [("activite_%d" % i) for i in indices] + ["activite_divers"]
    total = sum(_montant(p + "_montant") for p in lignes_act)

    ligne = 0
    for prefixe in lignes_act:
        designation = (r2.get(prefixe + "_designation", "") or "").strip()
        code_act = (r2.get(prefixe + "_code", "") or "").strip()
        montant = _montant(prefixe + "_montant")
        if not (designation or code_act or montant):
            continue
        ligne += 1
        if prefixe == "activite_divers" and not designation:
            designation = "Divers"
        if designation:
            variables[("NO_FR2B_1", ligne)] = (designation, T)
        if code_act:
            variables[("NO_FR2B_2", ligne)] = (code_act, T)
        if montant:
            variables[("NO_FR2B_3", ligne)] = (montant, N)
            if total:
                variables[("NO_FR2B_4", ligne)] = (round(montant / total, 4), N)
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

    # Les notes litteraires viennent en premier : si un code est a la fois
    # couvert ici et par une correspondance d'identification, c'est la fiche
    # qui gagne (aucun chevauchement aujourd'hui, mais l'ordre est explicite).
    notes_texte.update(fixes)
    return notes_texte, variables


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
