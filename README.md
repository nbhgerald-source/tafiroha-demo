# TAFIROHA en ligne — version de DÉMONSTRATION

⚠️ Ceci est une **copie de démo** du projet réel, créée pour permettre des
tests en ligne sans toucher aux vraies données clients. Le projet de
référence reste `tafiroha_app/` (inchangé). Les seules différences avec le
projet réel sont :

- `seed_demo.py` : crée un client fictif, un compte client de démo et une
  balance comptable fictive (mais équilibrée) sur 2025, pour qu'il y ait
  immédiatement des données à explorer (BILAN, RESULTAT, TFT, notes annexes).
- `wsgi_demo.py` : point d'entrée utilisé par le serveur de production
  (gunicorn) — initialise la base et charge le jeu de données de démo au
  démarrage (`app.py` lui-même n'est pas modifié).
- `Procfile`, `render.yaml`, `gunicorn` dans `requirements.txt` : nécessaires
  pour l'hébergement en ligne.

## Comptes de test — et comment cacher les mots de passe

Les emails sont fixes (`admin@tafiroha.local` pour le cabinet,
`client@demo.local` pour le client de démo), mais **les mots de passe ne
sont jamais écrits dans le code ni dans ce fichier** — ce dépôt étant
potentiellement public sur GitHub, un mot de passe écrit ici serait visible
par n'importe qui. À la place :

- Sur Render : aller dans le service → **Environment** → ajouter deux
  variables `ADMIN_PASSWORD` et `DEMO_PASSWORD` avec les mots de passe de
  ton choix, puis redéployer (bouton "Manual Deploy"). C'est la méthode
  recommandée — ces variables ne sont jamais commitées sur GitHub.
- Si tu ne définis pas ces variables, l'appli génère un mot de passe
  aléatoire à chaque démarrage et l'affiche **uniquement dans les logs
  privés de Render** (onglet "Logs" du service, lignes
  `ADMIN_PASSWORD non défini : mot de passe admin généré -> ...`) —
  personne d'autre que toi (connecté à ton compte Render) ne peut les voir.

Le compte admin voit tous les clients (et peut en créer d'autres) ; le
compte client de démo ne voit que sa propre page avec son historique
(exercice 2025 déjà chargé).

⚠️ Pour le dépôt GitHub : s'il était déjà créé en **public** sans cette
protection, considère que le mot de passe par défaut d'origine
(`admin1234`) a pu être visible dans l'historique Git — change-le
immédiatement via `ADMIN_PASSWORD` comme indiqué ci-dessus. Pour plus de
sûreté, tu peux aussi rendre le dépôt **privé** (Settings → Danger Zone →
Change visibility sur GitHub) ; Render continue de fonctionner avec un
dépôt privé.

## ⚠️ Important : données non persistantes sur l'hébergement gratuit

Render (plan gratuit) utilise un disque éphémère : à chaque redéploiement ou
redémarrage du service, le fichier `tafiroha.db` repart de zéro et le jeu de
données de démo est recréé automatiquement. C'est voulu pour un usage "test
comme un jeu" — ne sert pas à stocker de vraies données clients. Pour une
vraie mise en production, voir la section Sécurité du projet réel
(`tafiroha_app/README.md`) et prévoir un disque persistant ou une base
externe.

## Mise en ligne (GitHub + Render, gratuit)

### 1. Mettre ce dossier sur GitHub

Depuis ce dossier (`tafiroha_demo/`) :

```
git init
git add .
git commit -m "Démo TAFIROHA en ligne"
```

Puis sur https://github.com, créer un nouveau dépôt (vide, sans README), et
suivre les instructions affichées pour le lier et pousser le code, en
remplaçant `TON_COMPTE` et `NOM_DU_DEPOT` :

```
git remote add origin https://github.com/TON_COMPTE/NOM_DU_DEPOT.git
git branch -M main
git push -u origin main
```

### 2. Déployer sur Render

1. Aller sur https://render.com et créer un compte (gratuit, via GitHub
   directement c'est le plus simple).
2. "New +" → "Blueprint" (ou "Web Service" si l'option Blueprint n'apparaît
   pas) → choisir le dépôt GitHub créé à l'étape 1.
3. Render détecte automatiquement `render.yaml` et propose la configuration
   (nom `tafiroha-demo`, build/start command déjà renseignés). Valider.
4. Au bout de quelques minutes, Render donne une URL du type
   `https://tafiroha-demo.onrender.com` — c'est l'adresse à partager pour les
   tests.
5. Avant de partager le lien : va dans **Environment** et définis
   `ADMIN_PASSWORD` / `DEMO_PASSWORD` (voir section "Comptes de test"
   ci-dessus), puis "Manual Deploy" pour appliquer.

(Si l'option "Blueprint" n'est pas proposée par ton compte Render : choisir
"Web Service" classique, sélectionner le dépôt, et renseigner manuellement :
Build Command = `pip install -r requirements.txt`, Start Command =
`gunicorn wsgi_demo:application --bind 0.0.0.0:$PORT`.)

### Mettre à jour la démo plus tard

Toute modification poussée sur la branche `main` de GitHub redéploie
automatiquement sur Render (quelques minutes).

## Test en local avant mise en ligne

```
pip install -r requirements.txt
python3 wsgi_demo.py    # initialise + affiche les comptes de démo
python3 app.py          # lance le serveur sur http://127.0.0.1:8000
```

---

Pour la documentation complète des fonctionnalités (fidélité aux formules
Excel, périmètre couvert, etc.), voir `tafiroha_app/README.md` dans le
projet réel — identique ici, hors les ajouts de démo listés ci-dessus.
