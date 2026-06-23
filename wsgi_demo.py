"""Point d'entrée WSGI pour la version de démo en ligne (Render/gunicorn).

`app.py` (inchangé, identique au projet réel) n'initialise la base et ne crée
le compte admin par défaut que lorsqu'il est lancé directement
(`if __name__ == "__main__":`). Sous gunicorn, ce bloc n'est jamais exécuté.
Ce petit fichier fait donc l'initialisation + le chargement du jeu de données
fictif avant d'exposer l'objet WSGI `application` que gunicorn doit servir.

Mots de passe admin / client de démo : ne JAMAIS les laisser en dur dans le
code (visible publiquement si le dépôt GitHub est public). Ils sont repris
depuis les variables d'environnement ADMIN_PASSWORD / DEMO_PASSWORD, définies
dans le tableau de bord Render (Settings → Environment) — jamais commitées
sur GitHub. Si elles ne sont pas définies, un mot de passe aléatoire est
généré à chaque démarrage et n'apparaît que dans les logs privés de Render.

Démarrage en production : gunicorn wsgi_demo:application
"""
import os
import secrets

import db
import seed_demo

db.init_db()
seed_demo.run()

conn = db.get_conn()
try:
    admin_password = os.environ.get("ADMIN_PASSWORD") or secrets.token_urlsafe(9)
    conn.execute(
        "UPDATE users SET password_hash=? WHERE role='admin'",
        (db.hash_password(admin_password),),
    )
    demo_password = os.environ.get("DEMO_PASSWORD") or secrets.token_urlsafe(9)
    conn.execute(
        "UPDATE users SET password_hash=? WHERE email=?",
        (db.hash_password(demo_password), seed_demo.DEMO_USER_EMAIL),
    )
    conn.commit()
    if not os.environ.get("ADMIN_PASSWORD"):
        print("ADMIN_PASSWORD non défini : mot de passe admin généré -> %s" % admin_password)
    if not os.environ.get("DEMO_PASSWORD"):
        print("DEMO_PASSWORD non défini : mot de passe client de démo généré -> %s" % demo_password)
finally:
    conn.close()

from app import application  # noqa: E402,F401  (import après l'init, volontaire)
