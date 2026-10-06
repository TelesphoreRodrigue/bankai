# Application patrimoine — version 7 proposée

Cette version est une copie indépendante de l'interface Streamlit v6, de son cœur métier, de ses ressources cartographiques et du classeur fourni. Elle ne modifie pas l'application d'origine. Elle ne remplace pas l'interface Dash.

## Lancement local sous Windows

Depuis PowerShell, dans ce dossier :

```powershell
.\Lancer_v7.ps1
```

Le lanceur demande de choisir un identifiant et un mot de passe pour cette session locale. Ils ne sont pas écrits dans un fichier. Ouvrir ensuite http://127.0.0.1:8507 et saisir ces mêmes identifiants. Fermer le serveur avec Ctrl+C. Le lanceur utilise l'installation Python disponible et ne modifie pas ses dépendances.

Alternative : définir `MINFI_USERNAME` et `MINFI_PASSWORD`, puis exécuter `python -m streamlit run app7.py --server.address 127.0.0.1 --server.port 8507` dans ce dossier. Des secrets Streamlit `[auth]` avec `username` et `password` prennent priorité sur les variables d'environnement. Aucun identifiant de secours n'est fourni.

`requirements.txt` indique les versions de l'environnement effectivement testé (Python 3.11), et non une recommandation de déploiement. Sur une autre machine, installer ces dépendances dans un environnement Python séparé. Le dossier `assets` doit rester avec les fichiers Python.

## Données

Aucun classeur d'inventaire n'est fourni dans ce dépôt. Deux façons de charger les données :

- importer un classeur `.xlsx` depuis la barre latérale, rubrique « Source de données » : il reste en mémoire le temps de la session et n'est pas enregistré sur le serveur ;
- en local, placer le classeur dans ce dossier, à côté de `app7.py` : il est alors proposé par défaut.

Le classeur doit avoir la structure attendue (feuille `CENTRAL` et feuilles régionales, en-têtes de la fiche de recensement).

## Déploiement sur Streamlit Community Cloud

1. Sur https://share.streamlit.io, créer une application à partir de ce dépôt, branche `main`, fichier principal `app7.py`.
2. Dans les paramètres avancés, choisir Python 3.11 : les versions de `requirements.txt` ne sont pas publiées pour les versions plus récentes.
3. Toujours dans les paramètres avancés, renseigner les secrets :

   ```toml
   [auth]
   username = "identifiant_choisi"
   password = "mot_de_passe_long_et_unique"
   ```

Sans ces secrets, l'application affiche « Accès désactivé ». Ne jamais écrire d'identifiants dans le dépôt.

## Changements

- Exclusion de la ligne Excel entièrement vide avant toute imputation : 3 223 lignes et 4 910 biens, sans changement des montants.
- Source Excel chargée par son contenu : une modification du fichier invalide le cache même si son nom reste identique. Cache borné à huit entrées et une heure.
- Panneau de fiabilité sur le périmètre filtré : états imputés en nombre de biens et en pourcentage, indice sur états renseignés, couverture des coordonnées, quantités corrigées et VNC manquantes/reconstituées.
- Conservation de la règle métier v6 « état non reconnu ou absent = Bon état ». L'indice complémentaire rend visible sa sensibilité, sans prétendre représenter les biens non renseignés.
- Dispersion des points désactivée par défaut et activable explicitement avec avertissement. Sans dispersion, les repères identiques se superposent ; il ne s'agit pas encore d'une carte à bulles agrégées.
- Une règle d'audit en erreur interrompt le traitement avec son code, au lieu d'être omise silencieusement.
- VNC exposée globale comptée une seule fois par ligne d'inventaire touchée. Les sommes par type d'anomalie restent non additives entre elles.
- Reconstruction d'une VNC uniquement si acquisition et amortissement sont tous deux renseignés ; absence d'amortissement différente de zéro.
- Authentification sans mot de passe intégré ; prise en charge d'identifiants accentués.
- Nomenclatures exportées calculées sur le même périmètre filtré que la base ; marque de VNC imputée visible dans l'exploration.

## Limites

L'authentification reste un mécanisme local à compte partagé : pas de gestion des rôles, de limitation centralisée des tentatives, de journal des accès ou de SSO. Ce dossier n'est pas une certification pour mise en production. Les hypothèses métier, le rapprochement des fonds géographiques et les données manquantes nécessitent toujours une validation humaine. Les tests Streamlit vérifient l'exécution des vues, pas leur rendu pixel par pixel dans un navigateur.

Le classeur est une copie figée de celui examiné. Les mises à jour futures de l'original ne se propageront pas automatiquement à ce dossier.
