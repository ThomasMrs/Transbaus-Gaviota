# Transbaus-Gaviota

Application web locale pour trier des colis par baques, scanner des codes-barres et retrouver rapidement ou se trouve chaque colis.

## Fonctions

- ajout de baques depuis l'interface
- scan de code-barres avec la camera du navigateur
- saisie manuelle du code-barres
- suivi d'un colis par destination complete
- prise en compte optionnelle du numero destination de type `R447500`
- deplacement d'un colis d'une baque a une autre
- suppression de colis et de baques
- sauvegarde locale automatique dans le navigateur

## Lancer le site en local

Ouvrir un terminal PowerShell dans le dossier du projet puis lancer :

```powershell
.\serve-local.ps1
```

Ensuite ouvrir :

```text
http://localhost:4173
```

Le scan camera fonctionne mieux via `localhost` que via l'ouverture directe de `index.html`.
