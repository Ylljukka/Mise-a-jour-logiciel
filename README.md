Mise à jour des Logiciels Windows et des Packages Python

Vue d'ensemble
Ce script PowerShell permet de mettre à jour les logiciels installés sur les machines Windows en utilisant Winget, ainsi que les packages Python via pip. Il offre une interface graphique simple pour sélectionner les applications et packages à mettre à jour, et gère automatiquement le processus de mise à jour.


Fonctionnalités
Mise à jour des logiciels Windows : Rafraîchit et affiche une liste des logiciels installés pouvant être mis à jour via Winget.
Mise à jour des packages Python : Vérifie si pip est installé, met à jour pip, puis affiche et met à jour les packages Python obsolètes.
Interface utilisateur graphique (GUI) : Une interface utilisateur permet de sélectionner les logiciels et packages à mettre à jour.
Gestion des erreurs : Affiche des messages d'erreur en cas de problème lors des mises à jour.
Journalisation : Les messages relatifs aux mises à jour réussies ou échouées sont affichés dans une zone de texte dédiée.


Mises à jour des logiciels :

Le script affichera une fenêtre avec deux sections : une pour les logiciels Windows et une autre pour les packages Python.
Cochez les cases des logiciels et packages que vous souhaitez mettre à jour.
Cliquez sur le bouton "Mettre à jour" pour lancer la mise à jour.


Journalisation et messages :

Les messages concernant l'avancement des mises à jour et les erreurs éventuelles apparaîtront dans une zone de texte à droite de la fenêtre.


Personnalisation

Sélection de mise à jour : Vous pouvez personnaliser le script pour ajouter ou retirer certains logiciels ou packages de la liste de mise à jour.

Contribution
Les contributions sont les bienvenues ! Si vous trouvez des bugs ou souhaitez ajouter de nouvelles fonctionnalités, n'hésitez pas à créer une issue ou à soumettre une pull request.
