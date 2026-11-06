# -*- coding: utf-8 -*-
"""Export et import d'un dossier client complet (TAFIROHA).

Module autonome : il n'appelle aucune fonction de app.py et ne modifie
aucun comportement existant. Tout passe par les routes ajoutees dans
app.py, qui sont les seuls points de contact.

Un dossier exporte est une archive ZIP d'extension .tafiroha contenant :

  manifest.json  identite du dossier, comptages, empreinte SHA-256
  dossier.json   la totalite des donnees rattachees au client

But : transporter un dossier d'une installation a une autre
(bureau <-> en ligne, poste du client <-> poste du cabinet) pour
archivage, sauvegarde, reprise ou assistance a distance.

Ce qui voyage : le client, ses exercices, les balances N et N-1, le
detail TFT, toutes les saisies manuelles (chiffres et textes), la
selection du sommaire, la codification (regle client et exceptions par
exercice), les sens forces, les parametres de revision, le pointage et
les anomalies de revision, l'etat de depot des exercices.

Ce qui ne voyage pas : les comptes utilisateurs et leurs mots de passe,
qui sont propres a chaque installation. Les utilisateurs cites dans les
donnees (auteur d'une codification, auteur d'un pointage, deposant) sont
exportes sous forme d'adresse e-mail et rattaches a l'import au compte
local de meme e-mail s'il existe, sinon laisses vides.
"""

import datetime
import hashlib
import io
import json
import os
import zipfile

FORMAT = "tafiroha-dossier"
VERSION = 1
EXTENSION = ".tafiroha"

# Champs du client qui voyagent. Volontairement sans la colonne de
# proprietaire (owner_id en version bureau, created_by en ligne) : le
# proprietaire est l'utilisateur qui realise l'import.
CLIENT_CHAMPS = ("raison_sociale", "ncc", "ntd", "adresse", "created_at")

# Champs d'un exercice. depose_le / depose_par ne sont presents que sur
# les bases migrees, le module s'adapte.
EXERCICE_CHAMPS = ("annee", "date_debut", "date_fin", "libelle", "created_at", "depose_le")
EXERCICE_USERS = ("depose_par",)

# (table, colonnes referencant users)
TABLES_CLIENT = (
    ("client_comptes", ()),
    ("client_compte_mapping", ("cree_par",)),
    ("client_sens_override", ()),
    ("revision_parametres", ()),
)

TABLES_EXERCICE = (
    ("balance_lignes", ()),
    ("tft_detail_lignes", ()),
    ("note3_manuel", ()),
    ("note_texte", ()),
    ("sommaire_selection", ()),
    ("revision_pointage", ("pointe_par",)),
    ("revision_anomalies", ("traite_par",)),
    ("exercice_compte_mapping", ("cree_par",)),
)

LIBELLES = {
    "exercices": u"exercices",
    "balance_lignes": u"lignes de balance",
    "tft_detail_lignes": u"lignes de détail TFT",
    "note3_manuel": u"valeurs saisies",
    "note_texte": u"textes saisis",
    "sommaire_selection": u"choix du sommaire",
    "revision_pointage": u"pointages de la revue",
    "revision_anomalies": u"anomalies de révision",
    "exercice_compte_mapping": u"exceptions de codification",
    "client_comptes": u"comptes du client",
    "client_compte_mapping": u"codifications",
    "client_sens_override": u"sens forcés",
    "revision_parametres": u"paramètres de révision",
}


# --------------------------------------------------------------------
# Utilitaires de schema (tolerants aux bases non migrees)
# --------------------------------------------------------------------

def _table_existe(conn, table):
    r = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    return r is not None


def _colonnes(conn, table):
    return [r[1] for r in conn.execute("PRAGMA table_info(%s)" % table)]


def _colonne_proprietaire(conn):
    """owner_id en version bureau, created_by en version en ligne."""
    cols = _colonnes(conn, "clients")
    for c in ("owner_id", "created_by"):
        if c in cols:
            return c
    return None


def _email_de(conn, cache, uid):
    if uid is None:
        return None
    if uid not in cache:
        r = conn.execute("SELECT email FROM users WHERE id=?", (uid,)).fetchone()
        cache[uid] = r["email"] if r else None
    return cache[uid]


def _id_de_email(conn, cache, email):
    if not email:
        return None
    if email not in cache:
        r = conn.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone()
        cache[email] = r["id"] if r else None
    return cache[email]


def _nom_fichier(raison_sociale):
    base = "".join(
        c if (c.isalnum() or c in " -_") else "_" for c in (raison_sociale or "dossier")
    ).strip().replace(" ", "_")
    base = base[:60] or "dossier"
    return "%s_%s%s" % (base, datetime.date.today().isoformat(), EXTENSION)


# --------------------------------------------------------------------
# Export
# --------------------------------------------------------------------

def _exporter_lignes(conn, table, cle, valeur, user_cols, cache):
    cols = _colonnes(conn, table)
    rows = conn.execute(
        "SELECT * FROM %s WHERE %s=?" % (table, cle), (valeur,)
    ).fetchall()
    out = []
    for r in rows:
        d = {}
        for c in cols:
            if c == "id" or c == cle:
                continue
            if c in user_cols:
                d[c + "_email"] = _email_de(conn, cache, r[c])
            else:
                d[c] = r[c]
        out.append(d)
    return out


def construire(conn, client_id, origine="TAFIROHA"):
    """Rassemble tout le dossier d'un client dans un dictionnaire."""
    client = conn.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()
    if client is None:
        raise ValueError("Client introuvable")

    cache = {}
    cols_client = _colonnes(conn, "clients")
    data = {
        "format": FORMAT,
        "version": VERSION,
        "origine": origine,
        "exporte_le": datetime.datetime.now().replace(microsecond=0).isoformat(" "),
        "client": dict(
            (c, client[c]) for c in CLIENT_CHAMPS if c in cols_client
        ),
        "tables_client": {},
        "exercices": [],
        "tables_exercice": {},
    }

    for table, user_cols in TABLES_CLIENT:
        if not _table_existe(conn, table):
            continue
        data["tables_client"][table] = _exporter_lignes(
            conn, table, "client_id", client_id, user_cols, cache
        )

    cols_exo = _colonnes(conn, "exercices")
    exercices = conn.execute(
        "SELECT * FROM exercices WHERE client_id=? ORDER BY annee", (client_id,)
    ).fetchall()
    for table, _ in TABLES_EXERCICE:
        data["tables_exercice"][table] = []

    for rang, exo in enumerate(exercices):
        d = dict((c, exo[c]) for c in EXERCICE_CHAMPS if c in cols_exo)
        for c in EXERCICE_USERS:
            if c in cols_exo:
                d[c + "_email"] = _email_de(conn, cache, exo[c])
        data["exercices"].append(d)
        for table, user_cols in TABLES_EXERCICE:
            if not _table_existe(conn, table):
                continue
            for ligne in _exporter_lignes(
                conn, table, "exercice_id", exo["id"], user_cols, cache
            ):
                ligne["_exo"] = rang
                data["tables_exercice"][table].append(ligne)

    return data


def resume(data):
    """Comptages lisibles d'un dossier, pour le manifeste et l'ecran."""
    res = {"exercices": len(data.get("exercices", []))}
    for table, lignes in sorted(data.get("tables_client", {}).items()):
        if lignes:
            res[table] = len(lignes)
    for table, lignes in sorted(data.get("tables_exercice", {}).items()):
        if lignes:
            res[table] = len(lignes)
    return res


def exporter(conn, client_id, origine="TAFIROHA"):
    """Retourne (nom_fichier, octets) de l'archive du dossier."""
    data = construire(conn, client_id, origine)
    corps = json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True).encode("utf-8")
    manifest = {
        "format": FORMAT,
        "version": VERSION,
        "origine": origine,
        "exporte_le": data["exporte_le"],
        "client": data["client"].get("raison_sociale"),
        "ncc": data["client"].get("ncc"),
        "annees": [e.get("annee") for e in data["exercices"]],
        "resume": resume(data),
        "sha256": hashlib.sha256(corps).hexdigest(),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "manifest.json",
            json.dumps(manifest, ensure_ascii=False, indent=1).encode("utf-8"),
        )
        z.writestr("dossier.json", corps)
    return _nom_fichier(data["client"].get("raison_sociale")), buf.getvalue()


# --------------------------------------------------------------------
# Lecture d'une archive
# --------------------------------------------------------------------

def lire(octets):
    """Ouvre une archive et retourne (manifest, data).

    Leve ValueError avec un message en clair sur tout fichier invalide :
    l'archive vient de l'exterieur, rien n'est suppose correct.
    """
    try:
        z = zipfile.ZipFile(io.BytesIO(octets))
    except Exception:
        raise ValueError(
            "Ce fichier n'est pas une archive de dossier TAFIROHA (%s)." % EXTENSION
        )
    noms = z.namelist()
    if "dossier.json" not in noms:
        raise ValueError("Archive incomplete : dossier.json est absent.")
    try:
        data = json.loads(z.read("dossier.json").decode("utf-8"))
    except Exception:
        raise ValueError("Le contenu du dossier est illisible (JSON invalide).")
    if not isinstance(data, dict) or data.get("format") != FORMAT:
        raise ValueError("Ce fichier n'est pas un dossier TAFIROHA.")
    if int(data.get("version") or 0) > VERSION:
        raise ValueError(
            "Ce dossier a ete produit par une version plus recente de "
            "l'application (format %s). Mettez l'application a jour avant de "
            "l'importer." % data.get("version")
        )
    manifest = {}
    if "manifest.json" in noms:
        try:
            manifest = json.loads(z.read("manifest.json").decode("utf-8"))
        except Exception:
            manifest = {}
    attendu = manifest.get("sha256")
    if attendu:
        reel = hashlib.sha256(z.read("dossier.json")).hexdigest()
        if reel != attendu:
            raise ValueError(
                "L'empreinte du dossier ne correspond pas : le fichier a ete "
                "modifie ou endommage pendant le transport."
            )
    if not manifest:
        manifest = {
            "format": FORMAT,
            "version": data.get("version"),
            "origine": data.get("origine"),
            "exporte_le": data.get("exporte_le"),
            "client": (data.get("client") or {}).get("raison_sociale"),
            "ncc": (data.get("client") or {}).get("ncc"),
            "annees": [e.get("annee") for e in data.get("exercices") or []],
            "resume": resume(data),
        }
    return manifest, data


# --------------------------------------------------------------------
# Import
# --------------------------------------------------------------------

def _inserer(conn, table, cle, valeur, lignes, user_cols, cache, avertissements):
    """Insere des lignes en ne gardant que les colonnes connues localement."""
    if not lignes:
        return 0
    cols_locales = set(_colonnes(conn, table))
    ignorees = set()
    n = 0
    for ligne in lignes:
        champs = {cle: valeur}
        for c, v in ligne.items():
            if c == "_exo":
                continue
            if c.endswith("_email"):
                base = c[: -len("_email")]
                if base in user_cols and base in cols_locales:
                    champs[base] = _id_de_email(conn, cache, v)
                continue
            if c in cols_locales:
                champs[c] = v
            else:
                ignorees.add(c)
        noms = list(champs.keys())
        conn.execute(
            "INSERT OR REPLACE INTO %s (%s) VALUES (%s)"
            % (table, ", ".join(noms), ", ".join("?" for _ in noms)),
            [champs[k] for k in noms],
        )
        n += 1
    if ignorees:
        avertissements.append(
            "%s : champs inconnus de cette version, ignores (%s)."
            % (table, ", ".join(sorted(ignorees)))
        )
    return n


def _supprimer_dossier(conn, client_id):
    """Vide toutes les donnees d'un client sans supprimer le client ni ses
    rattachements d'utilisateurs."""
    exos = [
        r["id"]
        for r in conn.execute("SELECT id FROM exercices WHERE client_id=?", (client_id,))
    ]
    for table, _ in TABLES_EXERCICE:
        if not _table_existe(conn, table):
            continue
        for eid in exos:
            conn.execute("DELETE FROM %s WHERE exercice_id=?" % table, (eid,))
    for table, _ in TABLES_CLIENT:
        if _table_existe(conn, table):
            conn.execute("DELETE FROM %s WHERE client_id=?" % table, (client_id,))
    conn.execute("DELETE FROM exercices WHERE client_id=?", (client_id,))


def nom_libre(conn, raison_sociale):
    """Trouve un nom de client non utilise, en suffixant si besoin."""
    base = (raison_sociale or "Dossier importe").strip() or "Dossier importe"
    nom = base
    i = 2
    while conn.execute(
        "SELECT 1 FROM clients WHERE raison_sociale=?", (nom,)
    ).fetchone():
        nom = "%s (%d)" % (base, i)
        i += 1
        if i > 999:
            break
    return nom


def importer(conn, octets, user_id, mode="nouveau", client_cible=None, raison_sociale=None):
    """Importe une archive.

    mode "nouveau"     : cree un nouveau client (nom suffixe si deja pris).
    mode "remplacer"   : vide puis reconstruit le client client_cible.
                         Le client, ses utilisateurs rattaches et son
                         identifiant sont conserves : les acces en place
                         continuent de fonctionner.

    Retourne un dictionnaire {client_id, mode, nom, resume, avertissements}.
    L'operation est atomique : en cas d'erreur, rien n'est ecrit.
    """
    manifest, data = lire(octets)
    if mode not in ("nouveau", "remplacer"):
        raise ValueError("Mode d'import inconnu.")

    avertissements = []
    cache = {}
    cols_client = set(_colonnes(conn, "clients"))
    cols_exo = set(_colonnes(conn, "exercices"))
    proprio = _colonne_proprietaire(conn)
    infos = data.get("client") or {}

    if not conn.in_transaction:
        conn.execute("BEGIN")
    try:
        if mode == "remplacer":
            if not client_cible:
                raise ValueError("Aucun client cible n'a ete choisi.")
            cible = conn.execute(
                "SELECT * FROM clients WHERE id=?", (int(client_cible),)
            ).fetchone()
            if cible is None:
                raise ValueError("Le client a remplacer est introuvable.")
            client_id = cible["id"]
            nom = cible["raison_sociale"]
            _supprimer_dossier(conn, client_id)
            maj = {}
            for c in ("ncc", "ntd", "adresse"):
                if c in cols_client and infos.get(c):
                    maj[c] = infos[c]
            if maj:
                conn.execute(
                    "UPDATE clients SET %s WHERE id=?"
                    % ", ".join("%s=?" % c for c in maj),
                    list(maj.values()) + [client_id],
                )
        else:
            nom = nom_libre(conn, raison_sociale or infos.get("raison_sociale"))
            champs = {"raison_sociale": nom}
            for c in ("ncc", "ntd", "adresse", "created_at"):
                if c in cols_client and infos.get(c) is not None:
                    champs[c] = infos[c]
            if proprio:
                champs[proprio] = user_id
            noms = list(champs.keys())
            cur = conn.execute(
                "INSERT INTO clients (%s) VALUES (%s)"
                % (", ".join(noms), ", ".join("?" for _ in noms)),
                [champs[k] for k in noms],
            )
            client_id = cur.lastrowid

        compte = {}
        for table, user_cols in TABLES_CLIENT:
            if not _table_existe(conn, table):
                continue
            n = _inserer(
                conn, table, "client_id", client_id,
                (data.get("tables_client") or {}).get(table) or [],
                user_cols, cache, avertissements,
            )
            if n:
                compte[table] = n

        ids_exo = []
        for d in data.get("exercices") or []:
            champs = {"client_id": client_id}
            for c in EXERCICE_CHAMPS:
                if c in cols_exo and d.get(c) is not None:
                    champs[c] = d[c]
            for c in EXERCICE_USERS:
                if c in cols_exo:
                    uid = _id_de_email(conn, cache, d.get(c + "_email"))
                    if uid is not None:
                        champs[c] = uid
            if "annee" not in champs:
                raise ValueError("Un exercice du dossier n'a pas d'annee.")
            noms = list(champs.keys())
            cur = conn.execute(
                "INSERT INTO exercices (%s) VALUES (%s)"
                % (", ".join(noms), ", ".join("?" for _ in noms)),
                [champs[k] for k in noms],
            )
            ids_exo.append(cur.lastrowid)
        compte["exercices"] = len(ids_exo)

        for table, user_cols in TABLES_EXERCICE:
            if not _table_existe(conn, table):
                lignes = (data.get("tables_exercice") or {}).get(table) or []
                if lignes:
                    avertissements.append(
                        "%s : table absente de cette version, %d ligne(s) non "
                        "importee(s)." % (table, len(lignes))
                    )
                continue
            par_exo = {}
            for ligne in (data.get("tables_exercice") or {}).get(table) or []:
                par_exo.setdefault(ligne.get("_exo", 0), []).append(ligne)
            n = 0
            for rang, lignes in par_exo.items():
                if not isinstance(rang, int) or rang < 0 or rang >= len(ids_exo):
                    avertissements.append(
                        "%s : %d ligne(s) rattachee(s) a un exercice absent du "
                        "dossier, ignoree(s)." % (table, len(lignes))
                    )
                    continue
                n += _inserer(
                    conn, table, "exercice_id", ids_exo[rang], lignes,
                    user_cols, cache, avertissements,
                )
            if n:
                compte[table] = n

        conn.commit()
    except Exception:
        conn.rollback()
        raise

    return {
        "client_id": client_id,
        "mode": mode,
        "nom": nom,
        "manifest": manifest,
        "resume": compte,
        "avertissements": avertissements,
    }
