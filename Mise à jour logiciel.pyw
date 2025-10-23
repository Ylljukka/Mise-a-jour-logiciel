'''
Logiciel permettant de faire les mises à jours d'autres logiciels installés sur le poste.
Ce dernier met aussi à jour les paquets Python.
Une première version existe en Powershell, mais devant la difficulté sur l'armonie de l'interface graphique
et le multi-threading, une version Python est élaborée

Création :      15/05/2025      Thomas REYNAUD  v3.0.0
Modification    xx/xx/xxxx      
'''

# les modules
import json, os, sys, webbrowser
from powershell_utils import PowershellRun
from PyQt5.QtWidgets import(
QApplication, QWidget, QLabel, QProgressBar, QPushButton, QMessageBox, QListWidget, 
QListWidgetItem, QHBoxLayout, QMainWindow, QMenuBar, QPlainTextEdit
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QColor, QMovie, QIcon
import re
##############################################
##############################################
# Classe pour la fenêtre de chargement initial
##############################################
##############################################
class Fenetre_Chargement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestionnaire de mise à jour 3.0.0")
        self.setWindowIcon(QIcon('graph/update.png'))
        self.setFixedSize(400, 100)
        self.logiciels_temp = []
        self.packages_temp = []

        self.label = QLabel(self)
        self.label.setText("Chargement en cours...")
        self.label.setGeometry(130,20,300,20)
        self.label.setStyleSheet("font-weight: bold;")

        self.pbar = QProgressBar(self)
        self.pbar.move(50,50)
        self.pbar.resize(350, 20)
        self.pbar.setValue(10)

        # Lancement de l'initialisation
        self.initialisation = Initialisation()
        self.initialisation.progress.connect(self.pbar.setValue)
        self.initialisation.status.connect(self.label.setText)
        self.initialisation.liste_logiciel.connect(self.lister_logiciel_temp)
        self.initialisation.liste_package.connect(self.lister_package_temp)
        self.initialisation.initialisation_terminee_signal.connect(self.run_fenetre_principale)
        self.initialisation.start()

########################################################################
# Fonction pour fermer la fenêtre de chargement et lancer la  principale
########################################################################
    def run_fenetre_principale(self):
        self.close()
        self.fenetre_principale = Fenetre_Principale()
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
    liste_logiciel = pyqtSignal(list)
    liste_package = pyqtSignal(dict)
    initialisation_terminee_signal = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.liste_logiciels_remplis = []
        self.liste_python_remplis = {}

    def run(self):
        self.maj_winget()
        self.search_software()
        self.search_package()
        self.initialisation_terminee()

######################################################################
# Fonction pour l'execution du script Powershell de vérif + maj winget
######################################################################
    def maj_winget(self):
        self.status.emit("Mise à jour de Winget...")
        self.progress.emit(30)


        chemin_script = os.path.abspath("files/maj_Winget.ps1")
        commande = f"& '{chemin_script}'"
        resultat = PowershellRun.run_command(commande)
        
        if resultat and resultat.returncode == 0:
            self.status.emit("Winget mis à jour avec succès !")
            QThread.msleep(1000)

        else:
            self.status.emit("Échec exécution script Winget")


#######################################################
# Fonction de recherche des logiciels à mettre à jours
#######################################################
    def search_software(self):
        self.status.emit("Recherche des logiciels...")
        self.progress.emit(60)
        
        try:
            commande = "Get-WinGetPackage -Source winget | Where-Object IsUpdateAvailable | Select-Object -ExpandProperty Id"
            resultat = PowershellRun.run_command(commande)

# On mets à jour la liste_logiciels_remplis puis on emet la liste_logiciels_remplis dans la liste_logiciel qui
# appartient à la classe Initialisation
            if resultat and resultat.returncode == 0:
                lignes = resultat.stdout.strip().splitlines()
                self.liste_logiciels_remplis = [ligne.strip() for ligne in lignes if ligne.strip()]
                self.liste_logiciel.emit(self.liste_logiciels_remplis)
            else:
                self.status.emit("Erreur lors de la récupération des logiciels")
                QThread.msleep(1000)
                self.liste_logiciels_remplis = []

        except Exception as e:
            self.status.emit(f"Erreur système : {str(e)}")
            self.liste_logiciels_remplis = []


#############################################################
# Fonction de recherche des packages Python à mettre à jours
#############################################################
    def search_package(self):
        self.status.emit("Détection de pip...")
        self.progress.emit(80)

# On regarde si pip est installé
        commande = "Get-Command pip -ErrorAction SilentlyContinue"
        resultat = PowershellRun.run_command(commande)
        if not resultat or resultat.returncode != 0:
            self.status.emit("pip n'est pas installé")
            return

# Si installé on met à jour pip puis on cherche les packages à mettre à jour        
        else:
            commande = "python -m pip install --upgrade pip"
            resultat = PowershellRun.run_command(commande)
            if resultat and resultat.returncode == 0:
                self.status.emit("Pip mis à jour !")
                QThread.msleep(1000)

            self.status.emit("Recherche de package Python...")
            self.progress.emit(90)
            try :
                commande = 'pip list --outdated --format=json'
                resultat = PowershellRun.run_command(commande)

# On formate le retour pour la liste des packages dans le dictionnaire
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
                        self.status.emit("Erreur de parsing JSON")
                        self.status.emit(f"Erreur JSON : {e}")
                        self.liste_python_remplis = {}

                else:
                    self.status.emit("Erreur lors de la récupération des logiciels")
                    QThread.msleep(1000)
                    self.liste_python_remplis = {}

            except Exception as e:
                self.status.emit(f"Erreur système : {str(e)}")
                self.liste_python_remplis = {}


###############################################################
# Fonction de fin d'initialisation avec fermeture de la fenêtre
###############################################################
    def initialisation_terminee(self):
        self.status.emit("Chargement terminé !")
        self.progress.emit(100)
        QThread.msleep(1000)
        self.initialisation_terminee_signal.emit()
        

#----------------------------------------------------------------------------------------------------
#----------------- Fin de l'initialisation, on passe aux classes de mise à jour !!!------------------
#----------------------------------------------------------------------------------------------------

###################################
###################################
# Classe pour la fenêtre principale
###################################
###################################
class Fenetre_Principale(QMainWindow):
    def __init__(self):
        super().__init__()

# La fenêtre principales
        self.setWindowIcon(QIcon('graph/update.png'))
        self.setWindowTitle("Gestionnaire de mise à jour v3.0.0")
        self.setFixedSize(700, 600)

# Les boutons "tout cocher"
        self.bouton_decocher_logiciels = QPushButton("Tout décocher", self)
        self.bouton_decocher_logiciels.setGeometry(270, 70, 100, 25)
        self.bouton_decocher_logiciels.clicked.connect(self.decocher_tout_logiciels)
        self.bouton_decocher_packages = QPushButton("Tout décocher", self)
        self.bouton_decocher_packages.setGeometry(270, 380, 100, 25)
        self.bouton_decocher_packages.clicked.connect(self.decocher_tout_packages)

# Les labels et les listes à cocher
        self.label_software = QLabel(self)
        self.label_software.setText("Liste des logiciels à mettre à jour :")
        self.label_software.setGeometry(10, 40, 230, 30)
        self.liste_logiciels_widget = QListWidget(self)
        self.liste_logiciels_widget.setGeometry(10, 70, 250, 270)
        self.label_software.show()

        self.label_packages = QLabel("Liste des packages Python à mettre à jour :", self)
        self.label_packages.setGeometry(10, 350, 250, 30)
        self.liste_packages_widget = QListWidget(self)
        self.liste_packages_widget.setGeometry(10, 380, 250, 130)
        self.label_packages.show()
        self.liste_packages_widget.show()

# Les boutons        
        self.bouton_quitter = QPushButton("Quitter", self)
        self.bouton_quitter.setGeometry(50, 530, 150, 50)
        self.bouton_quitter.clicked.connect(self.quitter)   
        self.bouton_quitter.show()

        self.bouton_mettre_à_jour = QPushButton("Mettre à jour", self)
        self.bouton_mettre_à_jour.setGeometry(500, 530, 150, 50)
        self.bouton_mettre_à_jour.clicked.connect(self.choix)   
        self.bouton_mettre_à_jour.show()

# La barre de progression
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(390, 500, 290, 20)
        self.pbar.setValue(0)
        self.pbar.hide()


# La fenêtre contenant les logs
        self.zone_logs = QListWidget(self)
        self.zone_logs.setGeometry(390, 70, 300, 400)
        self.zone_logs.setStyleSheet("font-weight: bold;")


# Le menu cascade
        self.menu_cascade = QMenuBar(self)
        self.setMenuBar(self.menu_cascade)
        self.fichier = self.menu_cascade.addMenu("Fichier")
        self.about = self.menu_cascade.addMenu("A propos")
        self.fichier.addAction("Quitter", self.quitter)
        self.about.addAction("Notes de version", self.afficher_notes_version)
        self.about.addAction("Dons", lambda: self.ouverture_web(0))
        self.about.addAction("Linkedin", lambda: self.ouverture_web(1))
        self.about.addAction("Github", lambda: self.ouverture_web(2))


# On charge l'équivalent du CSS Python
        self.chargement_css()


# Fonction pour importer le fichier theme.qss pour gérer le design
    def chargement_css(self):
        try:
            with open("files/theme.qss", "r") as style_file:
                style = style_file.read()
                self.setStyleSheet(style)
        except FileNotFoundError:
            QMessageBox.warning(self, "Erreur", "Fichier theme.qss non trouvé")


# Fonction pour receuillir les logs de mise à jour
    def log_message(self, message, success):
        item = QListWidgetItem(message)
        if success is True:
            item.setForeground(QColor("green"))
        elif success is False:
            item.setForeground(QColor("red"))
        else :
            item.setForeground(QColor("white")) 
        self.zone_logs.addItem(item)

# Fonction pour afficher les notes de version
    def afficher_notes_version(self):
        try:
            self.notes = QPlainTextEdit()
            with open('files/Notes de version.txt', encoding="utf-8") as f:
                text = f.read()
            self.notes.setPlainText(text)
            self.notes.move(500, 300)
            self.notes.resize(500, 300)
            self.notes.setWindowTitle("Notes de version")
            self.notes.setWindowIcon(QIcon('graph/update.png'))
            self.notes.show() 
        except FileNotFoundError:
            QMessageBox.warning(self, "Erreur", "Fichier Notes de version.txt non trouvé")
            return

# Fonction pour gérer les différentes ouvertures de liens web
    def ouverture_web(self, web_link):
        if web_link ==0:
            webbrowser.open_new('https://www.paypal.com/paypalme/Thomas6598')
        elif web_link ==1 :
            webbrowser.open_new('https://www.linkedin.com/in/thomas-reynaud-2a8b02197/')

        elif web_link ==2:
            webbrowser.open_new('https://github.com/Ylljukka?tab=repositories')



#############################################################
# Fonction pour afficher la liste de logiciel à mettre à jour
#############################################################
    def lister_logiciel(self, liste_logiciels):
        self.liste_logiciels_widget.clear()

        if not liste_logiciels:
            item = QListWidgetItem("Aucun logiciel à mettre à jour")
            item.setFlags(Qt.NoItemFlags)
            self.liste_logiciels_widget.addItem(item)
            self.bouton_decocher_logiciels.setEnabled(False) 
        else:
            for logiciel in liste_logiciels:
                item = QListWidgetItem(logiciel)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
# On les coches par défaut                
                item.setCheckState(Qt.Checked)  
                self.liste_logiciels_widget.addItem(item)

######################################################################
# Fonction pour afficher la liste des packages Python à mettre à jour
######################################################################
    def lister_packages(self, dict_packages):
        self.liste_packages_widget.clear()

        if not dict_packages:
            item = QListWidgetItem("Aucun package Python à mettre à jour")
            item.setFlags(Qt.NoItemFlags)
            self.liste_packages_widget.addItem(item)
            self.bouton_decocher_packages.setEnabled(False)
        else: 
            for nom, versions in dict_packages.items():
                texte = f"{nom} : {versions['old_version']} → {versions['new_version']}"
                item = QListWidgetItem(texte)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.liste_packages_widget.addItem(item)


#########################################
# Fonction pour les boutons tout décocher
#########################################

    def decocher_tout_logiciels(self):
        for index in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(index)
            item.setCheckState(Qt.Unchecked)

    def decocher_tout_packages(self):
        for index in range(self.liste_packages_widget.count()):
            item = self.liste_packages_widget.item(index)
            item.setCheckState(Qt.Unchecked)

#####################################################
# Fonction pour le bouton de fermeture de la fenêtre
#####################################################
    def quitter(self):
        self.close()

####################################################################################
# Fonction pour récupérer les logiciels et packages sélectionné avant la mise à jour
####################################################################################
    def choix(self):
        self.bouton_mettre_à_jour.setEnabled(False) 
        self.bouton_decocher_logiciels.setEnabled(False)
        self.bouton_decocher_packages.setEnabled(False) 
        logiciels_choisis = []

# On regarde ce qui est coché, puis on ajoute à une liste que l'on transmet au Qthread 
        logiciels_choisis = [self.liste_logiciels_widget.item(i).text()
        for i in range(self.liste_logiciels_widget.count())
        if self.liste_logiciels_widget.item(i).checkState() == Qt.Checked]

        packages_choisis = [self.liste_packages_widget.item(i).text()
        for i in range(self.liste_packages_widget.count())
        if self.liste_packages_widget.item(i).checkState() == Qt.Checked]

        if not logiciels_choisis and not packages_choisis:
            QMessageBox.warning(self, "Erreur", "Veuillez cocher un élément")
            self.reenable_bouton_mettre_a_jour()
            return  
        
        self.thread_maj = Mise_à_Jour(logiciels_choisis, packages_choisis)
        self.thread_maj.retour.connect(self.log_message)
        self.thread_maj.fin_maj.connect(self.reenable_bouton_mettre_a_jour)
        self.thread_maj.progress.connect(self.pbar.setValue)
        self.thread_maj.demarrer_progression.connect(self.demarrer_progression)
        self.thread_maj.bloquer_cases.connect(self.bloquer_les_cases)
        self.thread_maj.debut_logiciel.connect(self.afficher_spinner)
        self.thread_maj.fin_logiciel.connect(self.afficher_ok)
        self.thread_maj.fin_package.connect(self.afficher_ok)
        self.thread_maj.vider_spinner.connect(self.supprimer_tous_les_spinners)
        self.thread_maj.start()

    def reenable_bouton_mettre_a_jour(self):
        self.bouton_mettre_à_jour.setEnabled(True)

    def demarrer_progression(self):
        self.pbar.setValue(0)
        self.pbar.show()

# Fonction pour afficher le spinner pendant la mise à jour
    def afficher_spinner(self, logiciel):
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            self.liste_logiciels_widget.setItemWidget(item, None)
    # On ajoute le GIF seulement sur le logiciel en cours
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            if item.text() == logiciel:
                spinner = QLabel()
                movie = QMovie("graph/ZJFD.gif")
                spinner.setMovie(movie)
                movie.start()
                self.liste_logiciels_widget.setItemWidget(item, spinner)
                break

# fonction pour l'affichage de l'icone si maj ok/ko
    def afficher_ok(self, logiciel, success):
        symbole = " ✅" if success else " ❌"

    # Pour les logiciels
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            if item.text().startswith(logiciel):
                texte_sans_symbole = item.text().split("✅")[0].split("❌")[0].strip()
                item.setText(f"{texte_sans_symbole} {symbole}")
                break

    # Pour les packages Python
        for i in range(self.liste_packages_widget.count()):
            item_python = self.liste_packages_widget.item(i)
            if item_python.text().split(":")[0].strip() == logiciel:
                texte_sans_symbole = item_python.text().split("✅")[0].split("❌")[0].strip()
                item_python.setText(f"{texte_sans_symbole} {symbole}")
                break


# Pour bloquer les cases à cocher lors de la maj :
    def bloquer_les_cases(self, bloquer):
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            flags = item.flags()
            if bloquer:
                item.setFlags(flags & ~Qt.ItemIsEnabled)
            else:
                item.setFlags(flags | Qt.ItemIsEnabled)
                
        for i in range(self.liste_packages_widget.count()):
            item_python = self.liste_packages_widget.item(i)
            flags_python = item_python.flags()
            if bloquer:
                item_python.setFlags(flags_python & ~Qt.ItemIsEnabled)
            else:
                item_python.setFlags(flags_python | Qt.ItemIsEnabled)


    def supprimer_tous_les_spinners(self):
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            self.liste_logiciels_widget.setItemWidget(item, None)

#########################################
#########################################
# Classe pour les actions de mise à jour
#########################################
#########################################
class Mise_à_Jour(QThread):
    progress = pyqtSignal(int)
    debut_logiciel = pyqtSignal(str)
    demarrer_progression = pyqtSignal()
    bloquer_cases = pyqtSignal(bool)
    retour = pyqtSignal(str, object)
    fin_maj = pyqtSignal()
    fin_logiciel = pyqtSignal(str, bool)
    fin_package = pyqtSignal(str, bool)
    vider_spinner = pyqtSignal()


    def __init__(self, logiciels, packages):
        super().__init__()
        self.logiciels = logiciels
        self.packages = packages

# Lancement des mises à jours
    def run(self):
        self.demarrer_progression.emit()
        self.bloquer_cases.emit(True) 
        total = len(self.logiciels) + len(self.packages)
        avancee = 0

        for logiciel in self.logiciels:
            success = False 
            try:
                self.retour.emit(f"Mise à jour de {logiciel}...", None)
                self.debut_logiciel.emit(logiciel) 
                commande = f"winget upgrade --id {logiciel} --silent"
                result = PowershellRun.run_command(commande)
                if result.returncode == 0:
                    success = True
                    self.retour.emit(f"{logiciel} mis à jour", True)
                else:
                    erreur = ''
                    if hasattr(result, 'stderr') and result.stderr:
                        erreur = self.nettoyage_log(result.stderr)
                    elif hasattr(result, 'stdout') and result.stdout:
                        erreur = self.nettoyage_log(result.stdout)
                    else:
                        erreur = "Erreur inconnue"
                        success = False
                    self.retour.emit(f"Échec maj {logiciel} : {erreur}", False)
            except Exception as e:
                self.retour.emit(f"Erreur {logiciel} : {e}", False)

            self.fin_logiciel.emit(logiciel, success)
            QApplication.processEvents()
            QThread.msleep(1000)
            avancee += 1
            self.progress.emit(int((avancee / total) * 100))

        for package in self.packages:
            success = False 
            nom = None
            try:
                self.retour.emit(f"Mise à jour de {package}...", None)
                self.debut_logiciel.emit(package) 
                nom = package.split(":")[0].strip()
                result = PowershellRun.run_command(f"pip install --upgrade {nom}")
                if result.returncode == 0:
                    success = True
                    self.retour.emit(f"{package} mis à jour", True)
                else:
                    erreur = ''
                    if hasattr(result, 'stderr') and result.stderr:
                        erreur = self.nettoyage_log(result.stderr)
                    elif hasattr(result, 'stdout') and result.stdout:
                        erreur = self.nettoyage_log(result.stdout)
                    else:
                        erreur = "Erreur inconnue"
                        success = False
                    self.retour.emit(f"Échec maj {package} : {erreur}", False)
            except Exception as e:
                self.retour.emit(f"Erreur {package} : {e}", False)

            self.fin_package.emit(nom, success)
            QApplication.processEvents()
            QThread.msleep(1000) 
            avancee += 1
            self.progress.emit(int((avancee / total) * 100))
        self.fin_maj.emit()
        self.vider_spinner.emit()


# Fonction pour nettoyer les logs d'échec de mise à jour
    def nettoyage_log(self, texte):
        texte = re.sub(r'[|\\/\-]', '', texte)
    #Supprimer les lignes qui ne contiennent que des espaces
        lignes = texte.splitlines()
        lignes_nettoyees = [ligne.strip() for ligne in lignes if ligne.strip()]
        return ' '.join(lignes_nettoyees)



    #---------------------------------------------------------------------------------------------
    #--------------------------------------Mon main ----------------------------------------------
    #---------------------------------------------------------------------------------------------

#Lancement de la phase d'initialisation
if __name__ == "__main__":
    app = QApplication(sys.argv)
    chargement = Fenetre_Chargement()
    chargement.show()
    sys.exit(app.exec_())
