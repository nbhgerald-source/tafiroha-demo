"""Couche base de données SQLite pour TAFIROHA en ligne."""
import sqlite3
import os
import hashlib
import binascii
import secrets

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "tafiroha.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    raison_sociale TEXT NOT NULL,
    ncc TEXT,
    ntd TEXT,
    adresse TEXT,
    created_by INTEGER,  -- id du gestionnaire créateur (NULL si créé par admin)
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('admin','gestionnaire','client')),
    client_id INTEGER REFERENCES clients(id),
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS exercices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id),
    annee INTEGER NOT NULL,
    date_debut TEXT,
    date_fin TEXT,
    libelle TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(client_id, annee)
);

CREATE TABLE IF NOT EXISTS balance_lignes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    periode TEXT NOT NULL CHECK(periode IN ('N','N1')),
    compte TEXT NOT NULL,
    designation TEXT,
    be_debit REAL DEFAULT 0,
    be_credit REAL DEFAULT 0,
    mvt_debit REAL DEFAULT 0,
    mvt_credit REAL DEFAULT 0,
    bs_debit REAL DEFAULT 0,
    bs_credit REAL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_balance_exercice ON balance_lignes(exercice_id, periode);

CREATE TABLE IF NOT EXISTS tft_detail_lignes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    periode TEXT NOT NULL CHECK(periode IN ('N','N1')),
    compte TEXT NOT NULL,
    designation TEXT,
    be_debit REAL DEFAULT 0,
    be_credit REAL DEFAULT 0,
    mvt_debit REAL DEFAULT 0,
    mvt_credit REAL DEFAULT 0,
    bs_debit REAL DEFAULT 0,
    bs_credit REAL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_tft_detail_exercice ON tft_detail_lignes(exercice_id, periode);

CREATE TABLE IF NOT EXISTS note3_manuel (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    sheet TEXT NOT NULL,
    coord TEXT NOT NULL,
    valeur REAL DEFAULT 0,
    UNIQUE(exercice_id, sheet, coord)
);

CREATE TABLE IF NOT EXISTS note_texte (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    sheet TEXT NOT NULL,
    champ TEXT NOT NULL,
    texte TEXT DEFAULT '',
    UNIQUE(exercice_id, sheet, champ)
);

CREATE TABLE IF NOT EXISTS sommaire_selection (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exercice_id INTEGER NOT NULL REFERENCES exercices(id),
    sheet TEXT NOT NULL,
    applicable INTEGER NOT NULL DEFAULT 1,
    UNIQUE(exercice_id, sheet)
);

CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT DEFAULT (datetime('now'))
);
"""


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _repair_dangling_users_old_refs(conn):
    """Filet de sécurité : si une future migration de schéma renomme un jour
    la table "users" (par ex. pour ajouter une contrainte CHECK), SQLite
    réécrit automatiquement les clauses REFERENCES des autres tables
    (clients.created_by, sessions.user_id) pour pointer vers le nom
    temporaire utilisé pendant le renommage. Si cette table temporaire est
    ensuite supprimée sans avoir pris soin de désactiver ce comportement
    (PRAGMA legacy_alter_table=ON), ces tables se retrouvent avec une
    référence vers une table qui n'existe plus — voir le même bug corrigé
    dans tafiroha_app/db.py le 2026-06-24. Pour l'instant, le schéma de
    tafiroha_demo inclut "gestionnaire" depuis sa création et ne nécessite
    aucun renommage, donc cette fonction est un no-op ; elle s'auto-répare
    si jamais une migration future réintroduit ce problème."""
    broken = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND sql LIKE '%users_old%'"
    ).fetchall()
    if not broken:
        return
    conn.execute("PRAGMA legacy_alter_table=ON")
    conn.execute("PRAGMA foreign_keys=OFF")
    for r in broken:
        name = r["name"]
        col_list = ", ".join(c["name"] for c in conn.execute("PRAGMA table_info(%s)" % name).fetchall())
        conn.execute("ALTER TABLE %s RENAME TO %s_brk" % (name, name))
        conn.executescript(SCHEMA)  # recrée "name" avec la bonne définition (IF NOT EXISTS)
        conn.execute("INSERT INTO %s (%s) SELECT %s FROM %s_brk" % (name, col_list, col_list, name))
        conn.execute("DROP TABLE %s_brk" % name)
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA legacy_alter_table=OFF")
    conn.commit()


def init_db():
    conn = get_conn()
    conn.executescript(SCHEMA)
    conn.commit()
    _repair_dangling_users_old_refs(conn)
    cur = conn.execute("SELECT COUNT(*) c FROM users WHERE role='admin'")
    if cur.fetchone()["c"] == 0:
        create_user(conn, "admin@tafiroha.local", "admin1234", "admin", None)
        print("Compte admin par défaut créé : admin@tafiroha.local / admin1234")
    conn.commit()
    conn.close()


def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return binascii.hexlify(salt).decode() + "$" + binascii.hexlify(dk).decode()


def verify_password(password, stored):
    try:
        salt_hex, _ = stored.split("$")
    except ValueError:
        return False
    salt = binascii.unhexlify(salt_hex)
    return hash_password(password, salt) == stored


def create_user(conn, email, password, role, client_id):
    conn.execute(
        "INSERT INTO users (email, password_hash, role, client_id) VALUES (?,?,?,?)",
        (email, hash_password(password), role, client_id),
    )


def create_session(conn, user_id):
    token = secrets.token_urlsafe(32)
    conn.execute("INSERT INTO sessions (token, user_id) VALUES (?,?)", (token, user_id))
    conn.commit()
    return token


def get_user_by_session(conn, token):
    if not token:
        return None
    row = conn.execute(
        "SELECT u.* FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token=?",
        (token,),
    ).fetchone()
    return row


def delete_session(conn, token):
    conn.execute("DELETE FROM sessions WHERE token=?", (token,))
    conn.commit()


if __name__ == "__main__":
    init_db()
    print("Base initialisée :", DB_PATH)
