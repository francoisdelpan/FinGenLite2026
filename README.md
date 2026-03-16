# FinGen 2026 - Google Sheets Extractions

Automatisation Google Apps Script pour generer des extractions depuis un Google Sheet de finances personnelles.

Modele Google Sheet :
[Ouvrir le modele](https://docs.google.com/spreadsheets/d/1l6D14WdTFZlT3jIZhaNRIPcM7vYAzVriXrFItoQNbho/edit?usp=sharing)

## Fonctionnalites

- Menu `Extractions` directement dans Google Sheets
- Extraction `par Tags`
- Extraction `par Compte`
- Extraction `par Wallet`
- Choix de la periode :
  - par mois
  - par plage de dates
- Export au choix :
  - en PDF
  - en Sheet
  - ou les deux
- Conservation optionnelle de la feuille generee
- Export PDF dans un dossier Google Drive defini

## Comportement des extractions

### par Tags

Recupere toutes les depenses de la periode selectionnee puis les regroupe par tag.

Sortie :
- tableau `Tag / Total depenses`

Nom d'export :
- `AAAAMMDD_HHMMSS_EXT-BY-TAG`

### par Compte

Recupere toutes les depenses de la periode pour le compte selectionne, puis genere un listing chronologique.

Sortie :
- tableau `NAME / DATE / TAG / AMOUNT`

Nom d'export :
- `AAAAMMDD_HHMMSS_EXT-BY-ACC`

### par Wallet

Recupere toutes les depenses de la periode pour le wallet selectionne, puis genere un listing chronologique.

Sortie :
- tableau `NAME / DATE / TAG / AMOUNT`

Nom d'export :
- `AAAAMMDD_HHMMSS_EXT-BY-WAL`

## Structure du projet

- [Code.gs](/Users/francoisdelpan/Documents/FinGen2026/Code.gs) : configuration globale
- [UI.gs](/Users/francoisdelpan/Documents/FinGen2026/UI.gs) : menu et ouverture des modales
- [ExtractionService.gs](/Users/francoisdelpan/Documents/FinGen2026/ExtractionService.gs) : logique metier des extractions
- [ExportService.gs](/Users/francoisdelpan/Documents/FinGen2026/ExportService.gs) : creation des feuilles et export PDF
- [Utils.gs](/Users/francoisdelpan/Documents/FinGen2026/Utils.gs) : fonctions utilitaires
- [ExtractionModal.html](/Users/francoisdelpan/Documents/FinGen2026/ExtractionModal.html) : interface utilisateur

Google Apps Script charge naturellement tous les fichiers `.gs` du projet.

## Installation

1. Cree une copie du modele Google Sheet.
2. Ouvre `Extensions > Apps Script`.
3. Copie les fichiers `.gs` et `ExtractionModal.html` dans le projet Apps Script.
4. Recharge le spreadsheet.
5. Le menu `Extractions` apparaitra dans la barre de menu.

## Configuration

La configuration principale se trouve dans [Code.gs](/Users/francoisdelpan/Documents/FinGen2026/Code.gs), dans l'objet `CFG`.

### Feuilles attendues

- `Transactions`
- `Dropdown`

### Colonnes a verifier

Dans `CFG`, adapte si besoin :

- `COL_DATE`
- `COL_NAME`
- `COL_AMOUNT`
- `COL_TAG`
- `COL_ACCOUNT`
- `COL_WALLET`
- `COL_DROPDOWN_TAG`
- `COL_DROPDOWN_ACCOUNT`
- `COL_DROPDOWN_WALLET`

### Dossier Drive pour les PDF

Deux modes sont possibles :

- Dossier fixe recommande :
  - `PDF_FOLDER_FALLBACK_TO_ROOT: false`
  - `PDF_FOLDER_ID: '...ID_DU_DOSSIER...'`
- Fallback automatique :
  - `PDF_FOLDER_FALLBACK_TO_ROOT: true`
  - le script tente d'enregistrer le PDF dans le meme dossier que le spreadsheet
  - si aucun dossier parent n'est trouve, le PDF est enregistre a la racine du Drive

## Personnalisation possible

Le projet est actuellement configure par code, ce qui est le plus simple et le plus fiable pour un deploiement Apps Script leger.

Une evolution possible serait d'ajouter une modale de configuration pour :

- choisir dynamiquement le dossier Drive des exports PDF
- enregistrer cet identifiant dans des proprietes Apps Script

Ce n'est pas indispensable pour utiliser le projet, mais cela peut etre utile si plusieurs utilisateurs doivent changer le dossier sans modifier le code.

## Notes

- Le script ne prend en compte que les depenses, c'est-a-dire les montants negatifs.
- Si un tag de transaction est vide ou inconnu, il est bascule dans `Autre / Null`.
- Pour `par Compte` et `par Wallet`, les transactions sont triees par date.
