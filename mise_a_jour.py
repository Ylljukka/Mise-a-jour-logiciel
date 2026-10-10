'''
Logiciel permettant de faire les mises à jours d'autres logiciels installés sur le poste.
Ce dernier met aussi à jour les paquets Python.
Une première version existe en Powershell, mais devant la difficulté sur l'armonie de l'interface graphique
et le multi-threading, une version Python est élaborée.

mise_à_jour.pyw est destiné au module qui effectue les mises à jours
Création :      15/05/2025      Thomas REYNAUD  v3.0.0
Modification    25/09/2026      Thomas REYNAUD  v3.1.0     

'''

import re, datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QThread, pyqtSignal
from commande_utils import Commande_Run


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
    resume = pyqtSignal(int, int)


    def __init__(self, logiciels, packages):
        super().__init__()
        self.logiciels = logiciels
        self.packages = packages
        self.annule = False
        self.reussites = 0
        self.echecs = 0

# Lancement des mises à jours
    def run(self):
        self.demarrer_progression.emit()
        self.bloquer_cases.emit(True) 
        total = len(self.logiciels) + len(self.packages)
        avancee = 0

        for logiciel in self.logiciels:
            if self.annule:
                self.retour.emit("Mise à jour annulée", None)
                break
            success = False 
            try:
                heure = datetime.datetime.now().strftime("%H:%M:%S")
                self.retour.emit(f"[{heure}] Mise à jour de {logiciel}...", None)
                self.debut_logiciel.emit(logiciel) 
                commande = ["winget", "upgrade", "--id", logiciel, "--silent",  "--include-unknown"]
                result = Commande_Run.run_command(commande)
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
            if success:
                self.reussites += 1
            else:
                self.echecs += 1
            QApplication.processEvents()
            QThread.msleep(1000)
            avancee += 1
            self.progress.emit(int((avancee / total) * 100))

        # Appel à la fonction de paralélisation
        if not self.annule:
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = {executor.submit(self.maj_package, pkg): pkg for pkg in self.packages}
                for future in as_completed(futures):
                    nom, success = future.result()
                    self.fin_package.emit(nom, success)
                    if success:
                        self.reussites += 1
                    else:
                        self.echecs += 1
                    QApplication.processEvents()
                    avancee += 1
                    self.progress.emit(int((avancee / total) * 100))
        else:
            self.retour.emit("Mise à jour des packages annulée", None)        
        self.resume.emit(self.reussites, self.echecs)
        self.fin_maj.emit()
        self.vider_spinner.emit()


# Fonction de paralélisation des MAJ packages Python
    def maj_package(self, package):
        success = False
        nom = None
        try:
            self.retour.emit(f"Mise à jour de {package}...", None)
            self.debut_logiciel.emit(package)
            nom = package.split(":")[0].strip()
            commande = ["python", "-m", "pip", "install", "--upgrade", "pip"] if nom == "pip" else ["pip", "install", "--upgrade", nom]
            result = Commande_Run.run_command(commande)
            if result.returncode == 0:
                success = True
                self.retour.emit(f"{package} mis à jour", True)
            else:
                erreur = self.nettoyage_log(result.stderr) if result.stderr else "Erreur inconnue"
                self.retour.emit(f"Échec maj {package} : {erreur}", False)
        except Exception as e:
            self.retour.emit(f"Erreur {package} : {e}", False)
        return nom, success

# Fonction pour nettoyer les logs d'échec de mise à jour
    def nettoyage_log(self, texte):
        texte = re.sub(r'[|\\/\-]', '', texte)
        #Supprimer les lignes qui ne contiennent que des espaces
        lignes = texte.splitlines()
        lignes_nettoyees = [ligne.strip() for ligne in lignes if ligne.strip()]
        return ' '.join(lignes_nettoyees)
    
    
# Fonction d'annulation de la mise à jours suivantes    
    def demander_annulation(self):
        self.annule = True