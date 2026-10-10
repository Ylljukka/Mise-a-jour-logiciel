'''
Module externe afin de pouvoir lancer des commandes systèmes (winget, pip, etc.)
directement, sans passer par PowerShell.
Cette classe est de type statique afin de pouvoir être appelée sans créer un objet de la classe
Ce module est de ce fait réutilisable

Création :      23/05/2025      Thomas REYNAUD  v1.0.0
Modification    03/10/2026      Thomas REYNAUD  v2.0.0 - suppression de la dépendance PowerShell
'''

import subprocess

class Commande_Run:
    @staticmethod
    # Fonction pour lancer une commande système (liste d'arguments)
    def run_command(args):
        try:
            # Constante pour ne pas afficher de fenêtre
            CREATE_NO_WINDOW = 0x08000000
            resultat = subprocess.run(
                args,
                creationflags=CREATE_NO_WINDOW,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace"
            )
            return resultat
        except Exception as e:
            return subprocess.CompletedProcess(args, returncode=-1, stdout="", stderr=f"Erreur système : {e}")