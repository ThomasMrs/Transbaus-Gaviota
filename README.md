# Transbaus-Gaviota

Application web de tri de colis pour TransBaus / Gaviota.

Le site permet de scanner un code-barres ou une etiquette complete, de ranger chaque colis dans une baque, puis de retrouver rapidement son emplacement, sa destination et ses informations principales.

## Fonctions principales

- ajout, renommage et suppression de baques
- scan de code-barres depuis la camera ou une image
- scan d'etiquette complete par photo avec OCR
- remplissage automatique des informations du colis
- gestion d'une destination complete, pas seulement du departement
- prise en compte optionnelle du numero destination de type `R447500`
- comparaison d'un bon de livraison PDF avec les colis enregistres
- recherche par code-barres, numero destination, destination, client, reference ou route
- deplacement d'un colis d'une baque a une autre
- suppression d'un colis
- sauvegarde des donnees metier dans Supabase pour un acces depuis n'importe quel appareil

## Informations gerees pour un colis

Le formulaire peut enregistrer :

- code-barres / numero de commande
- numero de commande exploitable pour la comparaison PDF
- numero destination
- destination complete
- client
- description
- route detaillee
- reference
- date
- poids
- numero de colis
- baque actuelle

## Fonctionnement conseille

1. Choisir la baque actuelle.
2. Utiliser `Scanner etiquette complete` pour prendre une photo de l'etiquette.
3. Verifier les informations remplies automatiquement.
4. Corriger si besoin, puis enregistrer le colis.
5. Utiliser la recherche ou le deplacement entre baques si le colis change d'emplacement.

Important :

- le `numero destination` suffit pour enregistrer un colis
- pour la comparaison avec un bon de livraison PDF, il faut aussi que le `numero de commande` soit reconnu ou saisi

## Configuration Supabase

Les cles publiques Supabase sont configurees dans [src/supabase-shared-state.mjs](src/supabase-shared-state.mjs).

Avant le premier usage, ouvrez l'editeur SQL de Supabase et executez :

```text
supabase/shared_state.sql
```

Cela cree la table `public.shared_state` et les policies minimales pour que l'application puisse lire et ecrire son etat partage.
Le meme script cree aussi une configuration d'acces privee et la fonction RPC `public.verify_site_access(text)` utilisee par l'ecran de connexion.

Important :

- le hash du mot de passe n'est plus stocke dans le frontend
- apres avoir execute le script, le login verifie le mot de passe cote Supabase
- le script migre automatiquement le hash actuellement utilise par l'application
- pour passer plus tard a un hash `bcrypt`, utilisez la commande commentee en bas de `supabase/shared_state.sql`

## Lancer le site

Depuis PowerShell, a la racine du projet :

```powershell
npm start
```

Le serveur local sert seulement les fichiers statiques.
Les donnees metier sont synchronisees directement avec Supabase.

Ouvrez ensuite :

```text
http://localhost:4173
```

Pour le bureau, utilisez l'adresse reseau du PC qui heberge le serveur :

```text
http://IP_DU_PC:4173
```

Pour un acces public, vous pouvez aussi publier le site statique sur GitHub Pages.
Dans les deux cas, les scans sont ecrits dans Supabase et deviennent visibles sur tous les appareils.

## Verifications locales

Le projet dispose maintenant d'une verification minimale :

```powershell
npm run check
```

Cela lance :

- `npm run lint` pour verifier le JavaScript modulaire
- `npm run test` pour verifier les parseurs et normalisations metier

## Mise en ligne sur GitHub Pages

Le depot contient un workflow GitHub Pages dans `.github/workflows/pages.yml`.

Fonctionnement :

- publication automatique a chaque `push` sur la branche `main`
- deploiement du site statique sans build frontend
- configuration prevue pour limiter les warnings lies a la migration Node 24 des GitHub Actions

Important :

- GitHub Pages convient maintenant au mode partage global
- la synchronisation des donnees passe par Supabase, pas par un serveur Node local

## Dependances cote navigateur

Le projet utilise :

- `html5-qrcode` pour le scan de code-barres
- `Tesseract.js` pour l'OCR des etiquettes

Important :

- `html5-qrcode` est embarque dans le depot via `vendor/html5-qrcode.min.js`
- `Tesseract.js` est charge depuis un CDN au chargement de la page
- sans connexion internet, l'OCR d'etiquette peut ne pas fonctionner si le script n'est pas deja charge

## Limites actuelles

- il n'y a plus de sauvegarde locale des scans: si Supabase ou Internet est coupe, les nouveaux scans ne sont pas conserves apres fermeture ou rechargement
- la qualite de l'OCR depend fortement de la nettete, de la lumiere et du cadrage de la photo
- certaines etiquettes peuvent necessiter une verification manuelle apres lecture automatique
- un colis enregistre sans numero de commande reste visible dans le site, mais il ne peut pas etre compare automatiquement avec un bon de livraison PDF
- les fichiers PDF restent stockes dans le navigateur qui les a importes, meme si leur analyse est visible dans l'etat partage
- le mot de passe d'entree est maintenant verifie cote Supabase, mais les donnees metier restent exposees via la cle publique tant que vous ne passez pas a une vraie authentification utilisateur

## Structure du projet

```text
.
|-- .github/workflows/pages.yml
|-- assets/
|-- supabase/
|   `-- shared_state.sql
|-- src/
|   |-- delivery-notes.mjs
|   |-- label-parser.mjs
|   |-- parcel-utils.mjs
|   |-- supabase-shared-state.mjs
|   `-- shared.mjs
|-- styles/
|   |-- components.css
|   |-- forms.css
|   |-- modals.css
|   |-- responsive.css
|   |-- shell.css
|   `-- tokens.css
|-- tests/
|-- vendor/
|-- app.js
|-- eslint.config.mjs
|-- index.html
|-- server.cjs
|-- serve-local.ps1
|-- styles.css
```
