'''
Logiciel permettant de faire les mises à jours d'autres logiciels installés sur le poste.
Ce dernier met aussi à jour les paquets Python.
Une première version existe en Powershell, mais devant la difficulté sur l'armonie de l'interface graphique
et le multi-threading, une version Python est élaborée.

main.pyw est destiné juste au lancement

Création :      15/05/2025      Thomas REYNAUD  v3.0.0
Modification    25/09/2026      Thomas REYNAUD  v3.1.0      
'''



import sys
from PyQt5.QtWidgets import QApplication
from chargement import Fenetre_Chargement

VERSION = "3.1.0"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    chargement = Fenetre_Chargement(VERSION)
    chargement.show()
    sys.exit(app.exec_())