'''
Logiciel permettant de faire les mises à jours d'autres logiciels installés sur le poste.
Ce dernier met aussi à jour les paquets Python.
Une première version existe en Powershell, mais devant la difficulté sur l'armonie de l'interface graphique
et le multi-threading, une version Python est élaborée.

chargement.pyw est destiné au chargement de l'application

Création :      15/05/2025      Thomas REYNAUD  v3.0.0
Modification    25/09/2026      Thomas REYNAUD  v3.1.0      
'''

import json, re, shutil, threading
from PyQt5.QtWidgets import QWidget, QLabel, QProgressBar
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from commande_utils import Commande_Run
from interface import Fenetre_Principale


##############################################
##############################################
# Classe pour la fenêtre de chargement initial
##############################################
##############################################
class Fenetre_Chargement(QWidget):
    def __init__(self, version):
        super().__init__()
        self.version = version
        self.setWindowTitle(f"Gestionnaire de mise à jour {version}")
        self.setWindowIcon(QIcon('graph/update.png'))
        self.setFixedSize(400, 130)
        self.logiciels_temp = []
        self.packages_temp = []

        self.label = QLabel(self)
        self.label.setText("Chargement en cours...")
        self.label.setGeometry(130,20,300,20)
        self.label.setStyleSheet("font-weight: bold;")
        
        #v3.1.0 : On autorise 2 labels à l'écran pour améliorer la perf vu qu'on paralélise
        self.label_python = QLabel(self)
        self.label_python.setText("")
        self.label_python.setGeometry(130, 40, 300, 20)
        self.label_python.setStyleSheet("font-weight: bold;")

        self.pbar = QProgressBar(self)
        self.pbar.move(50,65)
        self.pbar.resize(350, 20)
        self.pbar.setValue(10)

        # Lancement de l'initialisation
        self.initialisation = Initialisation()
        self.initialisation.progress.connect(self.pbar.setValue)
        self.initialisation.status.connect(self.label.setText)
        self.initialisation.status_python.connect(self.label_python.setText)
        self.initialisation.liste_logiciel.connect(self.lister_logiciel_temp)
        self.initialisation.liste_package.connect(self.lister_package_temp)
        self.initialisation.initialisation_terminee_signal.connect(self.run_fenetre_principale)
        self.initialisation.start()

########################################################################
# Fonction pour fermer la fenêtre de chargement et lancer la  principale
########################################################################
    def run_fenetre_principale(self):
        self.close()
        self.fenetre_principale = Fenetre_Principale(self.version)
        self.initialisation.liste_logiciel.connect(self.fenetre_principale.lister_logiciel)
        self.initialisation.liste_package.connect(self.fenetre_principale.lister_packages)
        self.fenetre_principale.lister_logiciel(self.logiciels_temp)
        self.fenetre_principale.lister_packages(self.packages_temp)
        self.fenetre_principale.show()

    def lister_logiciel_temp(self, logiciels):
        self.logiciels_temp = logiciels

    def lister_package_temp(self, packages):
        self.packages_temp = packages


###########################################
###########################################
# Classe pour les actions d'initialisation
###########################################
###########################################
class Initialisation(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    status_python = pyqtSignal(str)
    liste_logiciel = pyqtSignal(list)
    liste_package = pyqtSignal(dict)
    initialisation_terminee_signal = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.liste_logiciels_remplis = []
        self.liste_python_remplis = {}

#Fonction pour charger chaque étape (winget + pip)
    def run(self):
        self.etapes_terminees = 0

        thread_python = threading.Thread(target=self.search_package)
        thread_python.start()

        self.maj_winget()
        self.etape_terminee()

        self.search_software()
        self.etape_terminee()

        thread_python.join()
        self.etape_terminee()

        self.initialisation_terminee()

    def etape_terminee(self):
        self.etapes_terminees += 1
        self.progress.emit(int((self.etapes_terminees / 3) * 90))
        
######################################################################
# Fonction pour mettre à jour Winget
######################################################################
    def maj_winget(self):
        self.status.emit("Mise à jour de Winget...")
        args = ["winget", "upgrade", "--id", "Microsoft.AppInstaller", "--silent", "--accept-source-agreements", "--accept-package-agreements"]
        resultat = Commande_Run.run_command(args)
        code = resultat.returncode & 0xFFFFFFFF

        if code == 0:
            self.status.emit("Winget mis à jour avec succès !")
        elif code == 0x8A15002B:
            self.status.emit("Winget déjà à jour !")
        else:
            self.status.emit("Échec de la mise à jour de Winget")
        QThread.msleep(1000)
        
#######################################################
# Fonction de recherche des logiciels à mettre à jours
#######################################################
    def search_software(self):
        self.status.emit("Recherche des logiciels...")

        try:
            args = ["winget", "upgrade", "--include-unknown", "--accept-source-agreements"]
            resultat = Commande_Run.run_command(args)

            if resultat and resultat.returncode == 0:
                self.liste_logiciels_remplis = self.parser_sortie_winget(resultat.stdout)
                self.liste_logiciel.emit(self.liste_logiciels_remplis)
            else:
                self.status.emit("Erreur lors de la récupération des logiciels")
                QThread.msleep(1000)
                self.liste_logiciels_remplis = []

        except Exception as e:
            self.status.emit(f"Erreur système : {str(e)}")
            self.liste_logiciels_remplis = []

# Fonction pour extraire la colonne ID du tableau texte renvoyé par winget
    def parser_sortie_winget(self, texte):
        lignes = texte.splitlines()
        # Liste d'ID vide qu'on va remplir avec les logiciels
        ids = []
        position_id = None
        idx_version = None

        for ligne in lignes:
            ligne_maj = ligne.upper()
            if "ID" in ligne_maj and "VERSION" in ligne_maj and position_id is None:
                position_id = ligne_maj.index("ID")
                idx_version = ligne_maj.index("VERSION")
                continue

            #Les conditions pour pas ramener des morceaux de textes
            if position_id is None:
                continue

            if ligne.startswith("-"):
                continue

            # Ligne de résumé ("8 mises à niveau disponibles.") : on arrête tout
            if re.match(r"^\d+ ", ligne):
                break

            if len(ligne) <= position_id:
                continue

            id_logiciel = ligne[position_id:idx_version].strip()
            if id_logiciel:
                ids.append(id_logiciel)

        return ids

#############################################################
# Fonction de recherche des packages Python à mettre à jours
#############################################################
    def search_package(self):
        self.status_python.emit("Détection de pip...")

        # On regarde si pip est installé
        if shutil.which("pip") is None:
            self.status_python.emit("pip n'est pas installé")
            return

        # Si installé on met à jour pip puis on cherche les packages à mettre à jour
        else:
            args = ["python", "-m", "pip", "install", "--upgrade", "pip"]
            resultat = Commande_Run.run_command(args)
            if resultat and resultat.returncode == 0 and "already satisfied" in resultat.stdout:
                self.status_python.emit("Pip déjà à jour !")
                QThread.msleep(1000)
            elif resultat and resultat.returncode == 0:
                self.status_python.emit("Pip mis à jour !")
                QThread.msleep(1000)
            else:
                erreur = resultat.stderr.strip() if resultat and resultat.stderr else "Erreur inconnue"
                self.status_python.emit(f"Échec maj pip : {erreur}")
                print(erreur)
                QThread.msleep(2000)

            self.status_python.emit("Recherche de package Python...")
            try:
                args = ["pip", "list", "--outdated", "--format=json"]
                resultat = Commande_Run.run_command(args)

                if resultat and resultat.returncode == 0:
                    try:
                        data = json.loads(resultat.stdout)
                        self.liste_python_remplis = {
                            pkg['name']: {
                                'old_version': pkg['version'],
                                'new_version': pkg['latest_version']
                            } for pkg in data
                        }
                        self.liste_package.emit(self.liste_python_remplis)

                    except json.JSONDecodeError as e:
                        self.status_python.emit("Erreur de parsing JSON")
                        self.status_python.emit(f"Erreur JSON : {e}")
                        self.liste_python_remplis = {}

                else:
                    self.status_python.emit("Erreur lors de la récupération des packages Python")
                    QThread.msleep(1000)
                    self.liste_python_remplis = {}

            except Exception as e:
                self.status_python.emit(f"Erreur système : {str(e)}")
                self.liste_python_remplis = {}
                
                
###############################################################
# Fonction de fin d'initialisation avec fermeture de la fenêtre
###############################################################
    def initialisation_terminee(self):
        self.status.emit("Chargement terminé !")
        self.progress.emit(100)
        QThread.msleep(1000)
        self.initialisation_terminee_signal.emit()
        
