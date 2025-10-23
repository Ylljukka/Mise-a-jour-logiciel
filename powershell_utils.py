'''
Module externe afin de pouvoir lancer des commandes Powershell
juste en passant le paramètre de la commande.
Cette classe est de type statique afin de peut être appelée sans créer un objet de la classe
Ce module est de ce fait réutilisable

Création :      23/05/2025      Thomas REYNAUD  v1.0.0
Modification    xx/xx/xxxx      
'''

import subprocess

class PowershellRun:
    @staticmethod

# Fonction pour lancer une commande Powershell    
    def run_command(commande):
        try:
            # Constante pour ne pas afficher de fenêtre Powershell
            CREATE_NO_WINDOW = 0x08000000
            resultat = subprocess.run(
            [
            "powershell", "-ExecutionPolicy", "Bypass", 
            "-Command", commande
            ],
		    creationflags=CREATE_NO_WINDOW,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
            )
            return resultat
        except Exception as e:
            return f"Erreur système : {str(e)}"

