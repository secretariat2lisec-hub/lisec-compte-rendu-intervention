# LISEC - Mini application de compte rendu d'intervention

Ce dossier contient une premiere version complete de l'application mobile et du script Google.

## 1. Ce qui est prevu

- Application web mobile pour GitHub Pages.
- Utilisation sur tablette et telephone Android/Samsung.
- Aucun champ obligatoire.
- Niveaux fixes : Facade, RDC, R+1, R+2, R+3.
- Possibilite d'ajouter un niveau en plus.
- Possibilite d'ajouter autant de localisations que necessaire par niveau.
- Gravite : Faible, Moyenne, Forte, Critique.
- Photos depuis l'appareil photo ou la galerie.
- Compression des photos avant envoi.
- Recapitulatif avant envoi.
- Envoi vers Google Apps Script.
- Enregistrement dans Google Sheets.
- Stockage des photos dans Google Drive.
- Envoi d'un mail court a `secretariat2.lisec@gmail.com` et `monasspref@gmail.com`.
- Rapport Word joint au mail.
- Liens Drive des photos conserves dans Google Sheets et dans le mail.

## 2. Fichiers

- `index.html` : application mobile.
- `data.js` : listes modifiables.
- `manifest.json` : installation sur l'ecran d'accueil.
- `sw.js` : amelioration du comportement mobile.
- `Code.gs` : code a coller dans Google Apps Script.

## 3. Organisation Google Sheets conseillee

Je conseille un seul fichier Google Sheets : `LISEC - Comptes rendus interventions`.

Dans ce fichier, le script cree :

- un onglet `Interventions` pour la liste generale,
- un onglet `Observations` pour toutes les localisations,
- un onglet `Photos` pour tous les liens Drive,
- un onglet par intervention pour avoir une lecture rapide du compte rendu.

Une feuille par niveau serait possible, mais ce serait moins pratique a filtrer et a maintenir quand il y aura beaucoup d'interventions.

## 4. Rapport Word

Le plus fiable est de creer un modele Google Docs LISEC avec des champs comme :

- `{{DATE_VISITE}}`
- `{{INGENIEUR}}`
- `{{DESTINATAIRE}}`
- `{{ADRESSE_SITE}}`
- `{{DESCRIPTION_OUVRAGE}}`
- `{{CONSTRUCTION}}`
- `{{NOTE_VISITE}}`
- `{{CONCLUSION}}`
- `{{PRECONISATION}}`

Ensuite, dans `Code.gs`, il faudra mettre l'identifiant du modele dans :

```js
templateDocId: "",
```

Si ce champ reste vide, le script genere deja un rapport propre automatiquement.

## 5. Publication Google Apps Script

1. Aller sur https://script.google.com
2. Creer un nouveau projet.
3. Coller tout le contenu de `Code.gs`.
4. Cliquer sur `Deploy` puis `New deployment`.
5. Choisir `Web app`.
6. Executer en tant que : vous.
7. Acces : toute personne disposant du lien.
8. Copier l'URL qui se termine par `/exec`.
9. Ouvrir `index.html`.
10. Rechercher cette ligne :

```js
const GOOGLE_APPS_SCRIPT_URL = "";
```

11. Coller l'adresse entre les guillemets :

```js
const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/XXXXXXXX/exec";
```

L'adresse Google Apps Script ne doit pas etre visible pour les ingenieurs. Elle est configuree une fois dans le code avant la mise en ligne sur GitHub Pages.

## 6. Publication GitHub Pages

1. Creer un depot GitHub.
2. Mettre les fichiers `index.html`, `data.js`, `manifest.json`, `sw.js` et le dossier `assets`.
3. Activer GitHub Pages dans les reglages du depot.
4. Ouvrir le lien GitHub Pages sur le telephone ou la tablette.
5. Ajouter l'application a l'ecran d'accueil depuis le navigateur.

## 7. Point important pour beaucoup de photos

Les photos sont compressees avant l'envoi. Elles sont toujours stockees dans Google Drive.

Le mail essaie de joindre les photos, mais garde une limite de taille pour eviter un echec d'envoi. Si le mail devient trop lourd, les liens Drive restent disponibles dans le mail et dans le Google Sheet.
