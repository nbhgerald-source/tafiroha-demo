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
    # Unite de mesure de l'activite : la DGI n'accepte que les deux libelles
    # complets de sa liste (feuille Referentiels, K3:K4). Le formulaire de
    # l'application propose les abreviations "CA HT" / "VA" — il faut les
    # traduire, sinon la valeur est refusee comme hors liste.
    UNITES = {"CA HT": "Chiffre d'affaires HT (CA HT)",
              "VA": "Valeur ajoutée (VA)"}
    unite = str(r2.get("type_montant_activite", "") or "").strip()
    put("NO_FR2C_1_1", UNITES.get(unite, unite))

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


# ── Lignes de detail ajoutables ─────────────────────────────────────────────
#
# Les tableaux munis d'un bouton "+ Ajouter une ligne" n'ont pas un nombre de
# lignes connu d'avance : l'application range leurs valeurs sous des cles
# "<prefixe>_<colonne>_<rang>" et leurs libelles sous "<prefixe>_<rang>", et
# non sous une coordonnee de cellule. Le classeur, lui, attend ces lignes a
# partir d'une ligne fixe, donnee par la feuille "champs variables" du modele
# officiel. Sans la traduction ci-dessous, ces lignes n'existaient ni dans
# l'export XML ni dans l'export Excel : l'information etait perdue.
#
# Le classeur ne reserve qu'une ou deux lignes de detail par tableau : le
# total qui suit est une somme dynamique (INDIRECT/ADDRESS) prevue pour
# tolerer l'insertion de lignes. Quand la saisie en compte davantage — un
# inventaire d'immobilisations en contient couramment plusieurs dizaines — il
# faut donc inserer les lignes manquantes dans la feuille, et non ecraser le
# total et ce qui le suit.
#
#   (feuille, prefixe valeurs, prefixe libelle, colonne du libelle,
#    premiere ligne, lignes imposees, lignes reservees par le modele,
#    colonnes rangees en numerique)
#
# Le dernier element ne sert que lorsque les valeurs numeriques et les textes
# partagent le meme prefixe : la NOTE 13 range ses nombres sous "ASSOC_J_0" et
# ses textes sous "assoc_F_0". Ses colonnes E a I ont ete du numerique dans
# une version anterieure ; d'anciennes saisies peuvent subsister en base et ne
# doivent plus ressortir (un "584" fantome en N0 de compte contribuable).
TABLEAUX_DYNAMIQUES = (
    # FICHE R1 : domiciliations bancaires (zone ZW du classeur). Aucun code
    # DGI ne leur correspond : elles figurent dans le classeur, pas dans le
    # fichier XML, qui ne prevoit pas de champ pour les comptes bancaires.
    ("FICHE R1", "BANQUE", "banque_nom",   "J", 52,   None,         9, None),
    # NOTE 21 : produits accessoires a detailler par nature (tableau NOTE21B).
    ("NOTE 21",  "PA",     "pa_designation", "A", 28, None,         1, None),
    ("SUPPL1",   "PR",     "desig_pr",     "I", 13,   None,         52, None),
    # Les trois rubriques de l'etat annexe occupent des lignes fixes et non
    # contigues dans le classeur (D18/F18, D19/F19 puis D23/F23) : la DGI les
    # connait sous SUPPL1A1 pour les deux premieres et SUPPL1A2 pour la
    # troisieme. D'ou les lignes imposees plutot qu'une ligne de depart.
    ("SUPPL1",   "AX",     "desig_ax",     "A", None, (18, 19, 23), 0, None),
    ("SUPPL2",   "D",      "desig",        "A", 9,    None,         1, None),
    # SUPPL4 : colonne A = numero de compte, B = designation, F = date de
    # mise en service — exactement les colonnes du fichier d'inventaire
    # importe. La designation n'est donc pas en colonne A ici.
    ("SUPPL4",   "D",      "desig",        "B", 10,   None,         1, None),
    # NOTE 13 : la ligne 11 porte les "Apporteurs, capital non appele" et la
    # 12 le total : une seule ligne de detail est reservee.
    ("NOTE 13",  "assoc",  "assoc_nom",    "A", 10,   None,         1, ("J", "L")),
    ("NOTE 16C", "ACTIF",  "actif_desig",  "A", 11,   None,         1, None),
    ("NOTE 16C", "PASSIF", "passif_desig", "A", 14,   None,         1, None),
    # NOTE 30 et NOTE 3E : tableaux a lignes ajoutables eux aussi, comme le
    # confirment les boutons AjoutNOTE30B / NOTE30C / NOTE3E3 de la macro
    # officielle. Les lignes doivent etre rangees par feuille et par ligne
    # croissante : le decalage des tableaux suivants en depend.
    ("NOTE 30",  "CH",     "desig_ch",     "A", 10,   None,         1, None),
    ("NOTE 30",  "PR",     "desig_pr",     "A", 20,   None,         1, None),
    ("NOTE 3E",  "REEVAL", "reeval_poste", "A", 11,   None,         1, None),
    ("NOTE 32",  "D",      "desig",        "A", 11,   None,         2, None),
    ("NOTE 33",  "D",      "desig",        "A", 12,   None,         2, None),
)

# Colonnes saisies a part, en plus des colonnes de valeurs du tableau.
COLONNES_ANNEXES = {
    ("SUPPL1", "PR"): (("unite_pr", "J"),),           # unite de quantite choisie
    ("SUPPL4", "D"):  (("compte", "A"), ("date", "F")),  # compte et mise en service
    ("FICHE R1", "BANQUE"): (("banque_compte", "Q"),),   # numero de compte
}


_COORD = re.compile(r"^([A-Z]{1,3})([0-9]{1,5})$")


def _colonne_retenue(colonne, cols_nombre, numerique):
    """La colonne est-elle rangee dans le bon stockage ?

    Sans ce filtre, une saisie laissee par une version anterieure dans la
    mauvaise table ressortirait a l'export alors que l'ecran l'ignore."""
    if not cols_nombre:
        return True
    dans = colonne.upper() in tuple(c.upper() for c in cols_nombre)
    return dans if numerique else not dans


def _nb_lignes_saisies(note3, ntexte, feuille, prefixe, prefixes_libelle,
                       cols_nombre=None):
    """Nombre de lignes de detail reellement saisies sur un tableau."""
    maxi = -1
    motif = re.compile(r"^%s_([A-Za-z]{1,2})_(\d+)$" % re.escape(prefixe), re.I)
    for numerique, source in ((True, (note3 or {}).get(feuille) or {}),
                              (False, (ntexte or {}).get(feuille) or {})):
        for cle, valeur in source.items():
            if valeur in (None, ""):
                continue
            m = motif.match(str(cle))
            if m and _colonne_retenue(m.group(1), cols_nombre, numerique):
                maxi = max(maxi, int(m.group(2)))
    txt = (ntexte or {}).get(feuille) or {}
    for pref in prefixes_libelle:
        mot = re.compile(r"^%s_(\d+)$" % re.escape(pref))
        for cle, valeur in txt.items():
            if valeur in (None, ""):
                continue
            m = mot.match(str(cle))
            if m:
                maxi = max(maxi, int(m.group(1)))
    return maxi + 1


def plan_tableaux(note3, ntexte):
    """Etat de chaque tableau a lignes ajoutables pour cet exercice.

    Retourne une liste de dictionnaires :
        feuille, prefixe, prefixe_libelle, colonne_libelle, annexes,
        premiere_modele : ligne de depart dans le classeur d'origine,
        base            : ligne de depart apres insertion des lignes
                          supplementaires des tableaux situes au-dessus,
        imposees        : lignes fixes, pour les tableaux non extensibles,
        lignes          : nombre de lignes saisies,
        a_inserer       : nombre de lignes a inserer dans la feuille.
    """
    plan, decalage = [], {}
    for (feuille, prefixe, pref_lib, col_lib, premiere, imposees, reserve,
         cols_nombre) in TABLEAUX_DYNAMIQUES:
        annexes = COLONNES_ANNEXES.get((feuille, prefixe), ())
        n = _nb_lignes_saisies(
            note3, ntexte, feuille, prefixe,
            [pref_lib] + [p for p, _ in annexes], cols_nombre)
        d = decalage.get(feuille, 0)
        a_inserer = max(0, n - reserve) if (premiere and reserve) else 0
        plan.append({
            "feuille": feuille, "prefixe": prefixe,
            "prefixe_libelle": pref_lib, "colonne_libelle": col_lib,
            "annexes": annexes, "premiere_modele": premiere,
            "colonnes_nombre": cols_nombre,
            "base": (premiere + d) if premiere else None,
            "imposees": tuple(x + d for x in imposees) if imposees else None,
            "lignes": n, "a_inserer": a_inserer,
        })
        decalage[feuille] = d + a_inserer
    return plan


# Totaux que le classeur calcule par une formule auto-ajustable
# (SUM(INDIRECT(...))) : ces formules n'ont pas pu etre extraites dans
# sheets_raw.json, le moteur les voit donc comme vides. Sans cette reprise,
# les lignes de detail ajoutees ne remontent pas dans les totaux — la NOTE 21
# affichait un chiffre d'affaires ampute de ses produits accessoires.
#   feuille -> {cellule du total: (prefixe des valeurs, colonne)}
TOTAUX_DYNAMIQUES = {
    "NOTE 21": {"F27": ("PA", "F"), "G27": ("PA", "G")},
}


def totaux_dynamiques(note3):
    """Valeurs a ajouter aux saisies avant le calcul des feuilles."""
    import calc_engine as ce
    out = {}
    for feuille, cibles in TOTAUX_DYNAMIQUES.items():
        raw = (note3 or {}).get(feuille) or {}
        for cellule, (prefixe, colonne) in cibles.items():
            if cellule in raw:
                continue
            motif = re.compile(r"^%s_%s_(\d+)$" % (re.escape(prefixe), re.escape(colonne)))
            total = 0.0
            trouve = False
            for cle, valeur in raw.items():
                if motif.match(str(cle)):
                    total += ce._num(valeur)
                    trouve = True
            if trouve:
                out.setdefault(feuille, {})[cellule] = total
    return out


def saisies_pour_moteur(note3, feuilles=None):
    """Saisies a transmettre au moteur de calcul, totaux dynamiques inclus."""
    base = dict((f, dict(v)) for f, v in (note3 or {}).items()
                if feuilles is None or f in feuilles)
    for feuille, cellules in totaux_dynamiques(note3).items():
        if feuilles is None or feuille in feuilles:
            base.setdefault(feuille, {}).update(cellules)
    return base


def decalages_lignes(plan):
    """Lignes inserees par feuille : {feuille: [(ligne de depart, nombre)]}."""
    d = {}
    for e in plan or ():
        if e.get("a_inserer") and e.get("base"):
            depart = e["base"] + (e["lignes"] - e["a_inserer"])
            d.setdefault(e["feuille"], []).append((depart, e["a_inserer"]))
    return d


def ligne_decalee(ligne, insertions):
    """Position d'une ligne du modele apres insertion."""
    return int(ligne) + sum(n for dep, n in (insertions or ()) if int(ligne) >= dep)


# Colonnes dont e-impots impose la valeur : elle doit figurer telle quelle
# dans le referentiel officiel. L'apostrophe typographique, une majuscule ou
# un accent de trop suffisent a faire rejeter la liasse
# ("La valeur ne fait pas partie de celles autorisees").
#   (feuille, colonne) -> cle du referentiel
LISTES_IMPOSEES = {
    ("NOTE 13", "F"): "pays_residence",   # Nationalite
    ("NOTE 13", "G"): "pays_residence",   # Autres nationalites
    ("NOTE 13", "H"): "pays_residence",   # Pays de residence
    ("NOTE 13", "I"): "nature_actions",   # Ordinaires / Preferences
}


def _cle_comparaison(valeur):
    """Forme comparable d'un libelle : sans accent, sans casse, apostrophes
    et espaces unifies."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(valeur or ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    for apo in ("\u2019", "\u02bc", "`", "\u00b4"):
        s = s.replace(apo, "'")
    return " ".join(s.lower().split())


def _referentiels():
    try:
        import app
        return app.LISTES_NOTE13 or {}
    except Exception:
        return {}


def normaliser_listes(note3, ntexte):
    """Recale les valeurs saisies sur l'orthographe exacte du referentiel.

    Une valeur approchante (apostrophe courbe, accent manquant, casse
    differente) est remplacee par l'entree officielle. Une valeur qui n'y
    correspond pas est laissee telle quelle : c'est une erreur de saisie que
    le controle de revision signale."""
    listes = _referentiels()
    if not listes:
        return
    index = {}
    for cle, valeurs in listes.items():
        index[cle] = {}
        for v in valeurs or ():
            index[cle].setdefault(_cle_comparaison(v), v)
    for (feuille, colonne), cle in LISTES_IMPOSEES.items():
        officiel = index.get(cle)
        source = (ntexte or {}).get(feuille)
        if not officiel or not source:
            continue
        for coord in list(source):
            m = _COORD.match(str(coord))
            if not m or m.group(1) != colonne:
                continue
            exact = officiel.get(_cle_comparaison(source[coord]))
            if exact and exact != source[coord]:
                source[coord] = exact


def _ligne_dynamique(premiere, imposees, rang):
    """Ligne du classeur correspondant au rang d'une ligne de detail."""
    if imposees:
        return imposees[rang] if 0 <= rang < len(imposees) else None
    return premiere + rang


def _carte_cellules():
    """Correspondance codes DGI / cellules, chargee par l'application.

    Import tardif : app importe deja ce module, l'import en tete de fichier
    creerait un cycle."""
    try:
        import app
        return app.cellules_liasse()
    except Exception:
        return {}


def _pourcentages():
    """Cellules et colonnes du classeur au format pourcentage.

    Une saisie est toujours exprimee en points de pourcentage dans
    l'application : l'utilisateur tape 20 pour 20 %. Le classeur et la DGI
    attendent la fraction (0,2), comme le fait deja le pourcentage calcule de
    la FICHE R2. La conversion se fait donc au moment de l'export."""
    carte = _carte_cellules() or {}
    return (set(carte.get("pourcent_cellules") or ()),
            set(carte.get("pourcent_colonnes") or ()))


def _en_fraction(onglet, coord, valeur, cellules, colonnes):
    """Convertit une saisie en points de pourcentage vers une fraction."""
    colonne = "".join(c for c in str(coord) if c.isalpha())
    if ("%s!%s" % (onglet, coord)) not in cellules and \
       ("%s!%s" % (onglet, colonne)) not in colonnes:
        return valeur
    try:
        return float(valeur) / 100.0
    except (TypeError, ValueError):
        return valeur


def ce_num(v):
    import calc_engine as _ce
    return _ce._num(v)


def preparer_saisies(note3: dict, ntexte: dict, plan=None):
    """Rend les saisies exploitables par les deux exports.

    1. Les lignes de detail ajoutables recoivent leur vraie coordonnee de
       cellule, pour que l'export XML et l'export Excel les voient.
    2. Les taux saisis en points de pourcentage deviennent des fractions sur
       les cellules que le classeur presente au format pourcentage.

    plan : resultat de plan_tableaux(), pour que les lignes de detail soient
    placees aux memes coordonnees que dans le classeur etendu. Calcule ici si
    l'appelant ne le fournit pas.

    Les dictionnaires d'origine ne sont pas modifies : le moteur de calcul
    continue de lire les valeurs telles qu'elles sont en base (la NOTE 37
    applique D36/100 dans sa formule)."""
    cellules, colonnes = _pourcentages()
    n3 = dict((f, dict(v)) for f, v in (note3 or {}).items())
    nt = dict((f, dict(v)) for f, v in (ntexte or {}).items())

    # NOTE 37 D36 : le taux d'imposition est saisi sur le Sommaire de
    # l'application (le classeur, lui, lit =Sommaire!M43). Sans cette reprise,
    # la case "Taux" de la NOTE 37 part vide alors qu'elle est renseignee.
    taux = (nt.get("SOMMAIRE", {}) or {}).get("taux_bic_bnc", "")
    if str(taux).strip() and not n3.get("NOTE 37", {}).get("D36"):
        try:
            t = float(str(taux).replace("%", "").replace(",", ".").strip())
        except ValueError:
            t = 0.0
        if 0 < t < 1:
            t *= 100.0
        if t:
            n3.setdefault("NOTE 37", {})["D36"] = t

    # Les saisies deja rangees par coordonnee et situees sous un tableau
    # agrandi descendent avec lui (NOTE 13 : "Apporteurs, capital non
    # appele" et le total, places sous le detail des associes).
    plan = plan if plan is not None else plan_tableaux(note3, ntexte)
    for feuille, insertions in decalages_lignes(plan).items():
        for source in (n3.get(feuille), nt.get(feuille)):
            if not source:
                continue
            for cle in [c for c in source if _COORD.match(str(c))]:
                m = _COORD.match(str(cle))
                ligne = ligne_decalee(int(m.group(2)), insertions)
                if ligne != int(m.group(2)):
                    source["%s%d" % (m.group(1), ligne)] = source.pop(cle)

    for entree in plan:
        feuille = entree["feuille"]
        prefixe = entree["prefixe"]
        pref_lib = entree["prefixe_libelle"]
        col_lib = entree["colonne_libelle"]
        premiere = entree["base"]
        imposees = entree["imposees"]
        raw = n3.get(feuille, {})
        txt = nt.get(feuille, {})
        if not raw and not txt:
            continue
        # Casse du prefixe ignoree : la NOTE 13 range ses colonnes texte sous
        # "assoc_E_0" et ses colonnes numeriques sous "ASSOC_J_0".
        motif = re.compile(r"^%s_([A-Za-z]{1,2})_(\d+)$" % re.escape(prefixe), re.I)
        cols_nombre = entree.get("colonnes_nombre")
        for numerique, source, cible in ((True, raw, raw), (False, txt, txt)):
            for cle, valeur in list(source.items()):
                m = motif.match(str(cle))
                if not m or valeur in (None, ""):
                    continue
                if not _colonne_retenue(m.group(1), cols_nombre, numerique):
                    continue
                ligne = _ligne_dynamique(premiere, imposees, int(m.group(2)))
                if ligne:
                    cible["%s%d" % (m.group(1).upper(), ligne)] = valeur
        paires = [(pref_lib, col_lib)]
        paires.extend(entree["annexes"])
        for pref, col in paires:
            motif_lib = re.compile(r"^%s_(\d+)$" % re.escape(pref))
            for cle, valeur in list(txt.items()):
                m = motif_lib.match(str(cle))
                if not m or valeur in (None, ""):
                    continue
                ligne = _ligne_dynamique(premiere, imposees, int(m.group(1)))
                if ligne:
                    txt["%s%d" % (col, ligne)] = valeur
        if raw:
            n3[feuille] = raw
        if txt:
            nt[feuille] = txt

    # NOTE 13 : la colonne "Montant total" vaut valeur nominale x nombre de
    # titres, et le TOTAL ajoute la ligne "Apporteurs, capital non appele".
    # Ces deux valeurs sont calculees, jamais saisies : sans cette reprise
    # elles sortaient vides du classeur comme du fichier XML.
    entree13 = next((e for e in plan if e["feuille"] == "NOTE 13"), None)
    raw13 = n3.get("NOTE 13")
    if entree13 and raw13 and entree13["lignes"]:
        import calc_engine as ce13
        nominale = ce13._num(raw13.get("L8"))
        base13 = entree13["base"]
        total13 = 0.0
        for rang in range(entree13["lignes"]):
            ligne = base13 + rang
            nombre = ce13._num(raw13.get("J%d" % ligne))
            if nominale and nombre:
                raw13["K%d" % ligne] = nominale * nombre
            total13 += ce13._num(raw13.get("K%d" % ligne))
        ligne_apporteurs = base13 + entree13["lignes"]
        total13 += ce13._num(raw13.get("K%d" % ligne_apporteurs))
        if total13:
            raw13["K%d" % (ligne_apporteurs + 1)] = total13

    # Totaux que le classeur calcule par une formule auto-ajustable : ils
    # rejoignent les saisies, sans quoi la ligne "TOTAL PRODUITS ACCESSOIRES"
    # de la NOTE 21 partait vide dans le fichier XML.
    for feuille, cellules in totaux_dynamiques(note3).items():
        cible = n3.setdefault(feuille, {})
        for coord, valeur in cellules.items():
            cible.setdefault(coord, valeur)
        # Colonne de variation N/N-1, calculee comme le fait le classeur.
        if feuille == "NOTE 21":
            f27, g27 = ce_num(cible.get("F27")), ce_num(cible.get("G27"))
            if g27:
                cible.setdefault("H27", round((f27 - g27) / g27, 4))
            elif f27:
                cible.setdefault("H27", 1 if f27 > 0 else -1)

    normaliser_listes(n3, nt)

    if cellules or colonnes:
        for feuille, valeurs in n3.items():
            for coord in list(valeurs):
                if re.match(r"^[A-Z]{1,3}[0-9]{1,5}$", str(coord)):
                    valeurs[coord] = _en_fraction(
                        feuille, coord, valeurs[coord], cellules, colonnes)

    return n3, nt


# ── Lecture de la valeur d'une cellule ──────────────────────────────────────

def _get_value(computed: dict, note3: dict, ntexte: dict, onglet: str, row: int, col: int,
               directes=()):
    """
    Retourne (valeur, est_texte).

    - ACTIF  → coordonnée identique dans BILAN (col non décalée)
    - PASSIF → colonne +5 dans BILAN
    - Feuilles calculées → computed[onglet][coord]
    - Feuilles manuelles → note_texte d'abord (texte), puis note3_manuel (numérique)
      Si la cellule n'a jamais été saisie → (None, False) → sera ignorée
    """
    import calc_engine as ce

    coord = _num_to_col(col) + str(row)

    def _a_une_formule(feuille, c):
        """La cellule existe-t-elle dans le classeur d'origine ?

        Une feuille calculée contient aussi des cellules de SAISIE, sans
        formule (NOTE 3B, NOTE 4B...). La macro officielle n'exporte que les
        cellules dont le texte affiché n'est pas vide : une case de saisie
        laissée vide ne part pas. En renvoyant 0 pour toute coordonnée d'une
        feuille calculée, on envoyait un 0 dans des colonnes que la DGI
        attend sous forme de liste de valeurs — d'où « la valeur ne fait pas
        partie de celles autorisées »."""
        return ce.SHEETS_RAW.get(feuille, {}).get(c) is not None

    # Feuilles calculees directement en Python (COMP-CHARGES, NOTE 31,
    # NOTE 34) : leur cache fait foi cellule par cellule, le classeur n'ayant
    # pas de formule exploitable a comparer.
    if onglet in directes:
        v = computed.get(onglet, {}).get(coord)
        if v is not None:
            return v, False

    if onglet in ("ACTIF", "PASSIF"):
        c = _num_to_col(col + (5 if onglet == "PASSIF" else 0)) + str(row)
        if not _a_une_formule("BILAN", c):
            return None, False
        return computed.get("BILAN", {}).get(c, 0), False

    if onglet in computed:
        if _a_une_formule(onglet, coord):
            # Cellule avec formule : le 0 est une vraie valeur, on l'exporte.
            return computed[onglet].get(coord, 0), False
        # Cellule de saisie sur une feuille calculée : on bascule sur les
        # sources manuelles ci-dessous plutôt que de fabriquer un 0.

    if ntexte.get(onglet, {}).get(coord) not in (None, ""):
        return ntexte[onglet][coord], True

    if note3.get(onglet, {}).get(coord) is not None:
        return note3[onglet][coord], False

    return None, False  # cellule vide → à ignorer


# ── Formatage de la valeur (équivalent GetValeur VBA) ───────────────────────

def _format_value(v, is_text: bool):
    """
    Retourne la chaîne à écrire dans <valeur>, ou None si la cellule est vide.

    Pas d'apostrophe en préfixe. La fonction GetValeur du classeur officiel en
    ajoute une (`"'" & Format(valeur, "dd/mm/yyyy")`), mais elle écrit dans une
    CELLULE Excel : l'apostrophe y est un marqueur de saisie qui force le
    format texte et ne fait pas partie de la valeur. SaveAsXMLData exporte donc
    `01/01/2025`, pas `'01/01/2025`. La reprise littérale de l'apostrophe dans
    le XML faisait rejeter par e-impôts toutes les dates (« la valeur n'est pas
    une date valide ») et toutes les listes de codes (« la valeur ne fait pas
    partie de celles autorisées »).
    """
    if v is None:
        return None

    if is_text:
        s = str(v).strip()
        return s or None

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
        if s.startswith("0") and len(s) > 1:   # Code "01234" : zéro de tête à garder
            return s
        return str(int(n)) if n == int(n) else str(round(n, 4))
    except ValueError:
        return s or None


# ── Échappement XML ──────────────────────────────────────────────────────────

def _esc(s) -> str:
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


# ── Génération XML principale ────────────────────────────────────────────────

def generate_xml(conn, exercice_id: int, client: dict, exo: dict, periode: str = "N",
                 caches_python=None) -> bytes:
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
    manual_for_engine = saisies_pour_moteur(note3, set(other_calc))
    notes_cache = ce.compute_workbook(
        other_calc, balN, balN1,
        manual=manual_for_engine,
        row_range=(1, 200),
    ) if balN else {}

    # Les lignes de detail ajoutables prennent leur coordonnee de cellule et
    # les taux saisis passent en fraction. Volontairement APRES le calcul :
    # le moteur doit continuer de lire les valeurs telles qu'elles sont en
    # base (la NOTE 37 applique elle-meme D36/100 dans sa formule).
    directes = tuple((caches_python or {}).keys())

    plan = plan_tableaux(note3, ntexte)
    note3, ntexte = preparer_saisies(note3, ntexte, plan)
    # Les tableaux a lignes ajoutables peuvent en compter plus que la plage
    # reservee par le modele : c'est la saisie qui fixe leur hauteur, sans
    # quoi un inventaire de soixante immobilisations n'en exporterait qu'une.
    decal_feuilles = decalages_lignes(plan)
    plan_par_tableau = {}
    for e in plan:
        if e["premiere_modele"] and e["lignes"]:
            plan_par_tableau[(e["feuille"], e["premiere_modele"])] = e

    # Vue consolidée : computed[onglet][coord] = valeur
    computed: dict = dict(notes_cache)
    computed["BILAN"]    = bilan_cache
    computed["RESULTAT"] = resultat_cache
    computed["TFT"]      = tft_cache
    # Feuilles calculees en Python par l'application (COMP-CHARGES, NOTE 31,
    # NOTE 34) : elles rejoignent la vue consolidee, sans quoi ces trois
    # etats partaient vides sur e-impots.
    for feuille, cache in (caches_python or {}).items():
        computed[feuille] = dict(cache)

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
            # Une feuille dont un tableau a ete agrandi voit ses lignes
            # suivantes descendre : les saisies y ont ete replacees, la
            # lecture doit suivre le meme decalage.
            row = ligne_decalee(c["premiere_row"] + c["pos_ligne"] - 1,
                                decal_feuilles.get(c["onglet"]))
            col = c["premiere_col"] + c["pos_col"] - 1
            v, is_text = _get_value(computed, note3, ntexte, c["onglet"], row, col, directes)
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

        etendu = plan_par_tableau.get((cols[0]["onglet"], premiere_row))
        if etendu:
            premiere_row = etendu["base"]
            n_rows = max(n_rows, etendu["lignes"])

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
                elif etendu:
                    # Tableau a lignes ajoutables : ses lignes se lisent dans
                    # la saisie, pas dans la feuille calculee. Sans cela, la
                    # deuxieme ligne de detail renverrait la ligne de total
                    # qui la suit dans le modele.
                    col = c["premiere_col"] + c["pos_col"] - 1
                    coord = _num_to_col(col) + str(row)
                    txt = (ntexte.get(c["onglet"]) or {}).get(coord)
                    if txt not in (None, ""):
                        v, is_text = txt, True
                    else:
                        v, is_text = (note3.get(c["onglet"]) or {}).get(coord), False
                else:
                    col = c["premiere_col"] + c["pos_col"] - 1
                    v, is_text = _get_value(computed, note3, ntexte, c["onglet"], row, col, directes)
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
