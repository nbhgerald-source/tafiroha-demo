"""Point d'entrée WSGI pour la version de démo en ligne (Render/gunicorn).

`app.py` (inchangé, identique au projet réel) n'initialise la base et ne crée
le compte admin par défaut que lorsqu'il est lancé directement
(`if __name__ == "__main__":`). Sous gunicorn, ce bloc n'est jamais exécuté.
Ce petit fichier fait donc l'initialisation + le chargement du jeu de données
fictif avant d'exposer l'objet WSGI `application` que gunicorn doit servir.

Démarrage en production : gunicorn wsgi_demo:application
"""
import db
import seed_demo

db.init_db()
seed_demo.run()

from app import application  # noqa: E402,F401  (import après l'init, volontaire)
