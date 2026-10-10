# Gestionnaire de mise à jour

Programme Windows avec interface graphique (PyQt5) qui détecte les logiciels (via **winget**) et les packages Python (via **pip**) à mettre à jour. Vous cochez ce que vous voulez, et il lance les mises à jour.

Version actuelle : **3.1.0**. Les notes de version sont dans `files/Notes de version.txt`, et accessibles depuis le menu *À propos* du programme.

## Fonctionnalités

- Liste des logiciels et des packages Python à mettre à jour, avec cases à cocher
- Mise à jour des packages Python en parallèle
- Heure de début de chaque mise à jour dans les logs
- Bouton *Annuler* : l'élément en cours se termine, les suivants ne démarrent pas
- Résumé de fin (réussites et échecs) et export des logs en `.txt`
- Rafraîchissement de la liste depuis le menu *Fichier*
- Fonctionne avec un Windows en français ou en anglais

## Prérequis

- Windows 10 ou 11 avec **winget** (installé avec « App Installer »)
- Une connexion internet
- Pour la partie packages Python : Python et pip installés sur le poste. Sinon, cette partie est simplement ignorée.

## Installation

### Option 1 : l'exécutable (sans installer Python)

1. Téléchargez le dernier zip dans l'onglet [Releases](https://github.com/Ylljukka/Mise-a-jour-logiciel/releases).
2. Décompressez-le où vous voulez.
3. Lancez `MiseAJourLogiciel.exe`.

Ne séparez pas l'exe de ses dossiers `graph/` et `files/` : le programme en a besoin pour fonctionner. Déplacez toujours le dossier complet.

### Option 2 : depuis les sources

Avec [uv](https://docs.astral.sh/uv/) (Python 3.14 ou plus récent) :

```
git clone https://github.com/Ylljukka/Mise-a-jour-logiciel.git
cd Mise-a-jour-logiciel
uv sync
uv run main.pyw
```

Ou avec pip :

```
git clone https://github.com/Ylljukka/Mise-a-jour-logiciel.git
cd Mise-a-jour-logiciel
pip install -r requirements.txt
python main.pyw
```

Lancez toujours le programme depuis le dossier du projet, car les chemins vers `graph/` et `files/` en dépendent.

## Lancement automatique (tâche planifiée)

Dans le Planificateur de tâches Windows, renseignez le champ **« Commencer dans »** avec le dossier du programme. S'il est vide, Windows lance la tâche depuis un autre dossier et le programme ne trouve plus ses fichiers.

- Avec l'exe : programme `MiseAJourLogiciel.exe`, « Commencer dans » = son dossier.
- Avec les sources : programme `.venv\Scripts\pythonw.exe`, argument `main.pyw`, « Commencer dans » = dossier du projet.

## À savoir

- Certains logiciels demandent les droits administrateur : acceptez l'invite de Windows quand elle apparaît.
- Certains logiciels ne peuvent pas être mis à jour automatiquement (par exemple quand la technologie d'installation a changé). Winget demande alors de les désinstaller puis de les réinstaller.
- Le programme met à jour les packages Python du `pip` présent sur le poste.


## Auteur

Thomas REYNAUD