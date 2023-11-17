# Préparation d'un kit de boulonnerie

## Prérequis
Cet outil a besoin du fichier Excel contenant :
* un onglet décrivant le contenu des kits
* un onglet correspondant au modèle de kit de matériel
Par exemple, le fichier _Préparation Kits Boulonnerie U3_ (Récupéré).xlsm_ a servi d'exemple lors de la création de l'outil.

## Usage :
* Exécuter cristina.exe (sur Windows) ou cristina (sur macOS ou Linux)
* Choisir le fichier contenant les descriptions des kits
  Par exemple, le fichier _Préparation Kits Boulonnerie U3_ (Récupéré).xlsm_
* Choisir le module
* Choisir le cas de démontage
* Valider
  un fichier _impression.xls_ est généré. Il contient plusieurs onglets. Chaque onglet décrit les pièces d'un kit de démontage.
* Ouvrir le fichier _impression.xls_ et imprimer tous les onglets

Note : tous les cas de démontage ne sont pas possibles pour certains modules. Si vous choisissez une telle paire un message d'erreur vous indiquera que la sélection est vide ou erronée.

## Développement
L'outil a été développé sur Python 3.11
Le fichier requirement.txt contient les dépendances
Le fichier _build.sh_ permet d'assembler un exécutable sur macOS ou Linux
Le fichier _build.bat_ permet d'assembler un exécutable sur Windows
