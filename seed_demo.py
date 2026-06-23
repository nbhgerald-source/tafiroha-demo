"""Jeu de données fictif pour la version de démo en ligne de TAFIROHA en ligne.

Crée (de façon idempotente — ne duplique rien si on relance) :
  - un client fictif "ENTREPRISE DEMO SARL"
  - un compte d'accès client de démo (email/mot de passe ci-dessous)
  - un exercice 2025 avec une balance comptable N et N-1 entièrement fictive
    mais équilibrée (total débit = total crédit), pour que le BILAN, le
    COMPTE DE RESULTAT, le TFT et les notes annexes affichent de vraies
    valeurs calculées et soient utilisables pour les tests.

Ce script ne modifie en rien le projet réel ("tafiroha_app") : il ne vit que
dans cette copie de démo ("tafiroha_demo").
"""
import db

DEMO_CLIENT_NOM = "ENTREPRISE DEMO SARL"
DEMO_USER_EMAIL = "client@demo.local"
DEMO_USER_PASSWORD = "demo1234"

# (compte, designation, be_debit, be_credit, mvt_debit, mvt_credit, bs_debit, bs_credit)
BALANCE_N = [
    ("101000", "Capital social", 0, 10_000_000, 0, 0, 0, 10_000_000),
    ("106000", "Réserves", 0, 2_000_000, 0, 0, 0, 2_000_000),
    ("110000", "Report à nouveau", 0, 500_000, 0, 0, 0, 500_000),
    ("120000", "Résultat net (exercice précédent)", 0, 13_200_000, 0, 0, 0, 13_200_000),
    ("211000", "Terrains", 8_000_000, 0, 0, 0, 8_000_000, 0),
    ("218100", "Matériel et mobilier de bureau", 15_000_000, 0, 0, 0, 15_000_000, 0),
    ("281810", "Amortissements matériel et mobilier", 0, 6_000_000, 0, 1_500_000, 0, 7_500_000),
    ("311000", "Marchandises", 4_000_000, 0, 0, 0, 4_000_000, 0),
    ("401000", "Fournisseurs", 0, 3_000_000, 18_000_000, 19_000_000, 0, 4_000_000),
    ("411000", "Clients", 2_500_000, 0, 33_000_000, 32_000_000, 3_500_000, 0),
    ("421000", "Personnel, rémunérations dues", 0, 0, 1_800_000, 2_000_000, 0, 200_000),
    ("512000", "Banques", 5_000_000, 0, 32_000_000, 19_800_000, 17_200_000, 0),
    ("571000", "Caisse", 200_000, 0, 0, 0, 200_000, 0),
    ("601000", "Achats de marchandises", 0, 0, 19_000_000, 0, 19_000_000, 0),
    ("661000", "Charges de personnel", 0, 0, 2_000_000, 0, 2_000_000, 0),
    ("681000", "Dotations aux amortissements", 0, 0, 1_500_000, 0, 1_500_000, 0),
    ("701000", "Ventes de marchandises", 0, 0, 0, 33_000_000, 0, 33_000_000),
]

# Exercice précédent (N-1) : même structure, mise à l'échelle (x0.85) — une
# mise à l'échelle d'une balance équilibrée reste équilibrée.
def _scale(rows, factor):
    out = []
    for compte, designation, be_d, be_c, mvt_d, mvt_c, bs_d, bs_c in rows:
        out.append((
            compte, designation,
            round(be_d * factor), round(be_c * factor),
            round(mvt_d * factor), round(mvt_c * factor),
            round(bs_d * factor), round(bs_c * factor),
        ))
    return out

BALANCE_N1 = _scale(BALANCE_N, 0.85)


def run():
    conn = db.get_conn()
    try:
        client = conn.execute(
            "SELECT * FROM clients WHERE raison_sociale=?", (DEMO_CLIENT_NOM,)
        ).fetchone()
        if client is None:
            cur = conn.execute(
                "INSERT INTO clients (raison_sociale, ncc, ntd, adresse) VALUES (?,?,?,?)",
                (DEMO_CLIENT_NOM, "1234567A", "987654321", "Abidjan, Côte d'Ivoire (démo)"),
            )
            client_id = cur.lastrowid
            conn.commit()
            print("Client de démo créé :", DEMO_CLIENT_NOM)
        else:
            client_id = client["id"]

        user = conn.execute(
            "SELECT * FROM users WHERE email=?", (DEMO_USER_EMAIL,)
        ).fetchone()
        if user is None:
            db.create_user(conn, DEMO_USER_EMAIL, DEMO_USER_PASSWORD, "client", client_id)
            conn.commit()
            print("Compte client de démo créé :", DEMO_USER_EMAIL, "/", DEMO_USER_PASSWORD)

        exo = conn.execute(
            "SELECT * FROM exercices WHERE client_id=? AND annee=?", (client_id, 2025)
        ).fetchone()
        if exo is None:
            cur = conn.execute(
                "INSERT INTO exercices (client_id, annee, libelle, date_debut, date_fin) "
                "VALUES (?,?,?,?,?)",
                (client_id, 2025, "Exercice 2025 (démo)", "2025-01-01", "2025-12-31"),
            )
            exercice_id = cur.lastrowid
            conn.commit()
            print("Exercice de démo créé : 2025")

            for periode, rows in (("N", BALANCE_N), ("N1", BALANCE_N1)):
                for compte, designation, be_d, be_c, mvt_d, mvt_c, bs_d, bs_c in rows:
                    conn.execute(
                        "INSERT INTO balance_lignes "
                        "(exercice_id, periode, compte, designation, "
                        " be_debit, be_credit, mvt_debit, mvt_credit, bs_debit, bs_credit) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?)",
                        (exercice_id, periode, compte, designation, be_d, be_c, mvt_d, mvt_c, bs_d, bs_c),
                    )
            conn.commit()
            print("Balance de démo (N et N-1) chargée.")
    finally:
        conn.close()


if __name__ == "__main__":
    db.init_db()
    run()
