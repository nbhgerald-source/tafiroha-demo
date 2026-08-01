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
    is_default INTEGER NOT NULL DEFAULT 0,
    disabled INTEGER NOT NULL DEFAULT 0,
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
CREATE TABLE IF NOT EXISTS password_reset_tokens (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT DEFAULT (datetime('now')),
    expires_at TEXT NOT NULL,
    used INTEGER NOT NULL DEFAULT 0
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



def _migrate_default_accounts(conn):
    """Ajoute les colonnes/tables introduites avec les comptes par défaut."""
    cols = [r["name"] for r in conn.execute("PRAGMA table_info(users)").fetchall()]
    if "is_default" not in cols:
        conn.execute("ALTER TABLE users ADD COLUMN is_default INTEGER NOT NULL DEFAULT 0")
    if "disabled" not in cols:
        conn.execute("ALTER TABLE users ADD COLUMN disabled INTEGER NOT NULL DEFAULT 0")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS password_reset_tokens (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id),
            created_at TEXT DEFAULT (datetime('now')),
            expires_at TEXT NOT NULL,
            used INTEGER NOT NULL DEFAULT 0
        );
    """)
    conn.commit()

def _migrate_revision(conn):
    """Tables du module de révision assistée.

    Point critique : les justifications saisies par le collaborateur ne doivent
    JAMAIS être perdues quand on relance les contrôles après correction de la
    balance. Le rapprochement se fait sur la clé (exercice_id, code, compte),
    d'où l'index unique — une anomalie qui réapparaît à l'identique retrouve
    son traitement."""
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS revision_anomalies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exercice_id INTEGER NOT NULL REFERENCES exercices(id),
            code TEXT NOT NULL,
            famille TEXT NOT NULL,
            libelle TEXT NOT NULL,
            compte TEXT,
            montant REAL,
            severite TEXT NOT NULL,
            statut TEXT NOT NULL DEFAULT 'traiter',
            commentaire TEXT,
            traite_par INTEGER REFERENCES users(id),
            traite_le TEXT,
            detecte_le TEXT DEFAULT (datetime('now')),
            vu_le TEXT DEFAULT (datetime('now'))
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_revision_cle
            ON revision_anomalies(exercice_id, code, IFNULL(compte,''));
        CREATE INDEX IF NOT EXISTS idx_revision_exo
            ON revision_anomalies(exercice_id, severite);

        -- Mémoire des comptes rencontrés, par CLIENT et non par exercice :
        -- le paramétrage validé une fois se transmet aux exercices suivants.
        CREATE TABLE IF NOT EXISTS client_comptes (
            client_id INTEGER NOT NULL REFERENCES clients(id),
            cle INTEGER NOT NULL,
            compte_source TEXT,
            sens_resolu TEXT,
            niveau_resolution INTEGER,
            premier_exercice INTEGER,
            dernier_exercice INTEGER,
            valide INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (client_id, cle)
        );

        -- Comptes de régularisation : rattachement d'un compte non reconnu à
        -- un compte du plan, saisi par l'utilisateur. Clé sur le CLIENT et non
        -- l'exercice : le rattachement vaut pour tous les exercices suivants.
        CREATE TABLE IF NOT EXISTS client_compte_mapping (
            client_id INTEGER NOT NULL REFERENCES clients(id),
            cle_source INTEGER NOT NULL,
            cle_cible INTEGER NOT NULL,
            compte_source TEXT,
            motif TEXT,
            cree_par INTEGER REFERENCES users(id),
            cree_le TEXT DEFAULT (datetime('now')),
            PRIMARY KEY (client_id, cle_source)
        );

        CREATE TABLE IF NOT EXISTS client_sens_override (
            client_id INTEGER NOT NULL REFERENCES clients(id),
            cle INTEGER NOT NULL,
            sens TEXT NOT NULL,
            motif TEXT,
            PRIMARY KEY (client_id, cle)
        );

        CREATE TABLE IF NOT EXISTS revision_parametres (
            client_id INTEGER,
            cle TEXT NOT NULL,
            valeur TEXT NOT NULL
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_revision_param
            ON revision_parametres(IFNULL(client_id,0), cle);
    """)
    conn.commit()



def _migrate_user_clients(conn):
    """Rattachement d'un utilisateur a PLUSIEURS clients.

    users.client_id est conserve : il designe le client principal, celui vers
    lequel l'utilisateur est redirige apres connexion. La table de liaison
    porte l'ensemble des clients accessibles, client principal compris, ce qui
    permet d'ecrire un controle d'acces unique sans cas particulier.

    La reprise initiale recopie les rattachements existants : aucun utilisateur
    ne perd l'acces a son client lors de la mise a jour."""
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS user_clients (
            user_id   INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            client_id INTEGER NOT NULL REFERENCES clients(id),
            PRIMARY KEY (user_id, client_id)
        );
        CREATE INDEX IF NOT EXISTS idx_user_clients_user ON user_clients(user_id);
    """)
    conn.execute(
        "INSERT OR IGNORE INTO user_clients (user_id, client_id) "
        "SELECT id, client_id FROM users "
        "WHERE role='client' AND client_id IS NOT NULL"
    )
    conn.commit()

def init_db():
    conn = get_conn()
    conn.executescript(SCHEMA)
    conn.commit()
    _migrate_default_accounts(conn)
    _migrate_revision(conn)
    _migrate_user_clients(conn)
    _repair_dangling_users_old_refs(conn)
    # ── Comptes par défaut ───────────────────────────────────────────────────
    if not conn.execute("SELECT 1 FROM users WHERE email='admin@tafiroha.local'").fetchone():
        create_user(conn, "admin@tafiroha.local", "admin1234", "admin", None, is_default=1)
    if not conn.execute("SELECT 1 FROM users WHERE email='gestionnaire@demo.local'").fetchone():
        create_user(conn, "gestionnaire@demo.local", "gest1234", "gestionnaire", None, is_default=1)
    if not conn.execute("SELECT 1 FROM users WHERE email='client@demo.local'").fetchone():
        demo = conn.execute("SELECT id FROM clients WHERE raison_sociale='Client Démo'").fetchone()
        if not demo:
            conn.execute("INSERT INTO clients (raison_sociale) VALUES ('Client Démo')")
            conn.commit()
            demo = conn.execute("SELECT id FROM clients WHERE raison_sociale='Client Démo'").fetchone()
        create_user(conn, "client@demo.local", "demo1234", "client", demo["id"], is_default=1)
    # Compte demo public : pas de setup-account requis
    if not conn.execute("SELECT 1 FROM users WHERE email='clientdemo@tafiroha.com'").fetchone():
        demo = conn.execute("SELECT id FROM clients WHERE raison_sociale='Client Démo'").fetchone()
        if not demo:
            conn.execute("INSERT INTO clients (raison_sociale) VALUES ('Client Démo')")
            conn.commit()
            demo = conn.execute("SELECT id FROM clients WHERE raison_sociale='Client Démo'").fetchone()
        create_user(conn, "clientdemo@tafiroha.com", "demo2025", "client", demo["id"], is_default=0)
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


def create_user(conn, email, password, role, client_id, is_default=0):
    conn.execute(
        "INSERT INTO users (email, password_hash, role, client_id, is_default) VALUES (?,?,?,?,?)",
        (email, hash_password(password), role, client_id, is_default),
    )


def create_reset_token(conn, user_id):
    import datetime
    token = secrets.token_urlsafe(32)
    expires = (datetime.datetime.utcnow() + datetime.timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "INSERT INTO password_reset_tokens (token, user_id, expires_at) VALUES (?,?,?)",
        (token, user_id, expires),
    )
    conn.commit()
    return token


def get_valid_reset_token(conn, token):
    return conn.execute(
        "SELECT * FROM password_reset_tokens WHERE token=? AND used=0 AND expires_at > datetime('now')",
        (token,),
    ).fetchone()


def consume_reset_token(conn, token, new_password, user_id):
    conn.execute(
        "UPDATE users SET password_hash=?, is_default=0 WHERE id=?",
        (hash_password(new_password), user_id),
    )
    conn.execute("UPDATE password_reset_tokens SET used=1 WHERE token=?", (token,))
    conn.commit()


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
