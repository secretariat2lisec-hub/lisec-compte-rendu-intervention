# LISEC - Mini application de compte rendu d'intervention

Ce dossier contient une premiere version complete de l'application mobile et du script Google.

## 1. Ce qui est prevu

- Application web mobile pour GitHub Pages.
- Utilisation sur tablette et telephone Android/Samsung.
- Aucun champ obligatoire : le secretariat pourra completer les informations dans une phase ulterieure.
- Lieu de la visite separe en adresse, code postal et ville.
- Photo du lieu distincte, prise avec l'appareil photo ou ajoutee depuis la galerie.
- Client saisi dans un bloc libre comprenant le nom et l'adresse complete.
- Mission saisie dans un bloc libre.
- Quatre champs distincts pour les personnes presentes.
- Six champs distincts pour les elements de construction ; les champs vides sont ignores.
- Trois champs facultatifs pour les noms de diffusion.
- Niveaux fixes : Facade, RDC, R+1, R+2, R+3.
- Possibilite d'ajouter un niveau en plus.
- Possibilite d'ajouter autant de localisations que necessaire par niveau.
- Gravite : Faible, Moyenne, Forte, Critique.
- Photos depuis l'appareil photo ou la galerie.
- Compression des photos avant envoi.
- Deux parcours : envoi au secretariat pour complement ou generation directe du rapport.
- Recapitulatif adapte au parcours choisi.
- Avertissement avant generation directe lorsque des rubriques principales sont vides.
- Sauvegarde automatique du brouillon sur l'appareil.
- Envoi vers Google Apps Script.
- Enregistrement dans Google Sheets.
- Stockage des photos dans Google Drive.
- Envoi d'un mail court a `secretariat2.lisec@gmail.com` et `monasspref@gmail.com`.
- Rapport Word joint au mail uniquement pour le parcours de generation directe.
- Liens Drive des photos conserves dans Google Sheets et dans le mail.

## 2. Fichiers

- `index.html` : application mobile.
- `data.js` : listes modifiables.
- `manifest.json` : installation sur l'ecran d'accueil.
- `sw.js` : amelioration du comportement mobile.
- `Code.gs` : code a coller dans Google Apps Script.
- `Secretariat.html` : interface de controle a ajouter au projet Google Apps Script sous le nom `Secretariat`.

## 3. Organisation Google Sheets conseillee

Je conseille un seul fichier Google Sheets : `LISEC - Comptes rendus interventions`.

Dans ce fichier, le script cree :

- un onglet `Interventions` pour la liste generale,
- un onglet `Observations` pour toutes les localisations,
- un onglet `Photos` pour tous les liens Drive,
- un onglet `Dossiers` pour la file de travail du secretariat,
- un onglet par intervention pour avoir une lecture rapide du compte rendu.

Chaque intervention possede egalement un fichier `dossier.json` dans son dossier Drive. Il conserve les champs modifiables et les references des photos sans recopier les images.

Une feuille par niveau serait possible, mais ce serait moins pratique a filtrer et a maintenir quand il y aura beaucoup d'interventions.

## 4. Rapport Word

Le plus fiable est de creer un modele Google Docs LISEC avec des champs comme :

- `{{DATE_VISITE}}`
- `{{INGENIEUR}}`
- `{{DESTINATAIRE}}`
- `{{ADRESSE_SITE}}`
- `{{CODE_POSTAL}}`
- `{{VILLE}}`
- `{{CLIENT}}`
- `{{MISSION}}`
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
4. Dans le projet Apps Script, ajouter un fichier HTML nomme `Secretariat`.
5. Coller dans ce fichier tout le contenu de `Secretariat.html`.
6. Ouvrir les parametres du projet puis les proprietes du script.
7. Ajouter la propriete `SECRETARIAT_ACCESS_CODE` et choisir un code long reserve au secretariat.
8. Cliquer sur `Deploy` puis `New deployment`.
9. Choisir `Web app`.
10. Executer en tant que : vous.
11. Acces : toute personne disposant du lien, afin que l'application mobile puisse transmettre un dossier.
12. Copier l'URL qui se termine par `/exec`.
13. Ouvrir cette URL : elle affiche l'interface du secretariat et demande le code d'acces.
14. Ajouter `?api=status` a l'URL pour verifier uniquement l'etat du service.
La version fournie est deja reliee au deploiement Apps Script LISEC. Les etapes suivantes ne servent que si un nouveau deploiement cree une autre URL.

15. Ouvrir `index.html`.
16. Rechercher cette ligne :

```js
const GOOGLE_APPS_SCRIPT_URL = "";
```

17. Coller l'adresse entre les guillemets :

```js
const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/XXXXXXXX/exec";
```

L'adresse Apps Script sera presente dans le code public de l'application. La confidentialite des dossiers repose donc sur le controle du code `SECRETARIAT_ACCESS_CODE`, jamais sur le secret de l'URL.

Les fonctions de lecture, de modification et de generation refusent toute demande sans ce code. Le code est conserve uniquement en memoire dans la page du secretariat et doit etre ressaisi apres sa fermeture.

Documentation officielle :

- https://developers.google.com/apps-script/guides/web
- https://developers.google.com/apps-script/guides/html/communication

## 6. Publication GitHub Pages

1. Creer un depot GitHub.
2. Mettre les fichiers `index.html`, `data.js`, `manifest.json`, `sw.js` et le dossier `assets`.
3. Activer GitHub Pages dans les reglages du depot.
4. Ouvrir le lien GitHub Pages sur le telephone ou la tablette.
5. Ajouter l'application a l'ecran d'accueil depuis le navigateur.

## 7. Point important pour beaucoup de photos

Les photos sont compressees avant l'envoi. Elles sont toujours stockees dans Google Drive.

Le mail essaie de joindre les photos, mais garde une limite de taille pour eviter un echec d'envoi. Si le mail devient trop lourd, les liens Drive restent disponibles dans le mail et dans le Google Sheet.
