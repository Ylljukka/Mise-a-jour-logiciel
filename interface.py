'''
Logiciel permettant de faire les mises à jours d'autres logiciels installés sur le poste.
Ce dernier met aussi à jour les paquets Python.
Une première version existe en Powershell, mais devant la difficulté sur l'armonie de l'interface graphique
et le multi-threading, une version Python est élaborée.

interface.pyw est destiné à l'interface principale du logiciel

Création :      15/05/2025      Thomas REYNAUD  v3.0.0
Modification    25/09/2026      Thomas REYNAUD  v3.1.0     

'''

import datetime, webbrowser
from PyQt5.QtWidgets import (
    QFileDialog, QMainWindow, QLabel, QProgressBar, QPushButton, QMessageBox,
    QListWidget, QListWidgetItem, QMenuBar, QPlainTextEdit
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QMovie, QIcon
from mise_a_jour import Mise_à_Jour



class Fenetre_Principale(QMainWindow):
    def __init__(self, version):
        super().__init__()
        self.version = version

# La fenêtre principales
        self.setWindowIcon(QIcon('graph/update.png'))
        self.setWindowTitle(f"Gestionnaire de mise à jour v{version}")
        self.setFixedSize(800, 600)

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
        
        self.bouton_annuler = QPushButton("Annuler", self)
        self.bouton_annuler.setGeometry(280, 530, 150, 50)
        self.bouton_annuler.clicked.connect(self.annuler_maj)
        self.bouton_annuler.hide()

# La barre de progression
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(390, 480, 400, 20)
        self.pbar.setValue(0)
        self.pbar.hide()


# La fenêtre contenant les logs
        self.zone_logs = QListWidget(self)
        self.zone_logs.setGeometry(390, 70, 400, 400)
        self.zone_logs.setStyleSheet("font-weight: bold;")


# Le menu cascade
        self.menu_cascade = QMenuBar(self)
        self.setMenuBar(self.menu_cascade)
        self.fichier = self.menu_cascade.addMenu("Fichier")
        self.about = self.menu_cascade.addMenu("A propos")
        self.action_rafraichir = self.fichier.addAction("Rafraîchir", self.rafraichir)
        self.fichier.addAction("Exporter les logs", self.exporter_logs)
        self.action_quitter = self.fichier.addAction("Quitter", self.quitter)
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
    
# Fonction d'export de fichier        
    def exporter_logs(self):
        nom_defaut = f"logs_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        chemin, _ = QFileDialog.getSaveFileName(self, "Exporter les logs", nom_defaut, "Fichiers texte (*.txt)")
        if not chemin:
            return
        try:
            with open(chemin, "w", encoding="utf-8") as f:
                for i in range(self.zone_logs.count()):
                    f.write(self.zone_logs.item(i).text() + "\n")
            QMessageBox.information(self, "Export", f"Logs exportés vers {chemin}")
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Échec export : {e}")

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

# La fonction de rafraichissement
    def rafraichir(self):
        self.close()
        from chargement import Fenetre_Chargement
        self.nouvelle_fenetre_chargement = Fenetre_Chargement(self.version)
        self.nouvelle_fenetre_chargement.show()


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
        
        # On regarde ce qui est coché, puis on ajoute à une liste que l'on transmet au Qthread
        logiciels_choisis = [self.liste_logiciels_widget.item(i).text()
        for i in range(self.liste_logiciels_widget.count())
        if self.liste_logiciels_widget.item(i).checkState() == Qt.Checked]

        packages_choisis = [self.liste_packages_widget.item(i).text()
        for i in range(self.liste_packages_widget.count())
        if self.liste_packages_widget.item(i).checkState() == Qt.Checked]

        if not logiciels_choisis and not packages_choisis:
            QMessageBox.warning(self, "Erreur", "Veuillez cocher un élément")
            return

        self.bouton_mettre_à_jour.setEnabled(False) 
        self.bouton_decocher_logiciels.setEnabled(False)
        self.bouton_decocher_packages.setEnabled(False) 
        self.bouton_quitter.setEnabled(False)
        self.action_quitter.setEnabled(False)
        self.bouton_annuler.show()
        self.action_rafraichir.setEnabled(False)

        
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
        self.thread_maj.resume.connect(self.afficher_resume)
        self.thread_maj.start()

# Réactivation des boutons
    def reenable_bouton_mettre_a_jour(self):
        self.bouton_quitter.setEnabled(True)
        self.action_quitter.setEnabled(True)
        self.action_rafraichir.setEnabled(True)
        self.bouton_annuler.hide()
        
# Fonction d'annulation de la mise à jours        
    def annuler_maj(self):
        self.thread_maj.demander_annulation()
        self.bouton_annuler.setEnabled(False)
        QMessageBox.information(self, "Annulation", "Mise à jour annulée.\n L'élément en cours va terminer son installation.")

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


# Fonction d'affichage du résumé        
    def afficher_resume(self, reussites, echecs):
        QMessageBox.information(self, "Résumé", f"{reussites} réussite(s), {echecs} échec(s)")
        
        
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


#Fonction qui bloque la fermeture à la croix pendant le traitement
    def closeEvent(self, event):
        if hasattr(self, 'thread_maj') and self.thread_maj.isRunning():
            QMessageBox.warning(self, "Mise à jour en cours", "Impossible de fermer la fenêtre pendant une mise à jour.")
            event.ignore()
        else:
            event.accept()
            
        
    def supprimer_tous_les_spinners(self):
        for i in range(self.liste_logiciels_widget.count()):
            item = self.liste_logiciels_widget.item(i)
            self.liste_logiciels_widget.setItemWidget(item, None)