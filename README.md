# Copieur de fichiers Excel

Ce script Python vous permet de créer des copies d'un fichier Excel pour chaque jour d'un mois et d'une année spécifiés. Le script utilise la bibliothèque Tkinter pour créer une interface graphique (GUI) simple permettant de saisir l'année et le mois. Les copies sont enregistrées dans un répertoire portant le nom du mois et de l'année spécifiés.

## Prérequis

Avant d'exécuter le script, assurez-vous d'avoir installé les éléments suivants :

- Python (version 3 ou supérieure)
- La bibliothèque Pandas (peut être installée via `pip install pandas`)

## Utilisation

1. Ouvrez le script dans un environnement Python ou un éditeur de texte.

2. Modifiez la variable `source_file` pour spécifier le chemin d'accès à votre fichier Excel source. Le chemin d'accès par défaut est défini sur `'C:/Users/USERNAME/Documents/Projets/File-Multipli/source-file/testing.xlsx'`.

3. Exécutez le script à l'aide d'un interpréteur Python.

4. Une fenêtre intitulée "Copieur de fichiers Excel" apparaîtra.

5. Entrez l'année et le mois souhaités dans les champs de saisie respectifs.

6. Cliquez sur le bouton "Démarrer" pour lancer le processus.

7. Le script créera un répertoire portant le nom du mois et de l'année spécifiés (par exemple, "Juin 2023") s'il n'existe pas déjà.

8. Des copies du fichier source, portant le nom de la date (par exemple, "Rapport du 01 juin 2023.xlsx"), seront créées dans le répertoire de destination pour chaque jour du mois spécifié.

9. L'étiquette d'état en bas de la fenêtre affichera le chemin de chaque copie créée.

10. Une fois le processus terminé, vous pouvez fermer la fenêtre.

Note : Assurez-vous d'avoir les autorisations nécessaires pour lire le fichier source et écrire dans le répertoire de destination.

## Dépendances

Les bibliothèques suivantes sont utilisées dans ce script :

- `os` : Fournit des fonctions pour interagir avec le système d'exploitation.
- `shutil` : Offre des opérations de haut niveau pour la copie et la suppression de fichiers.
- `pandas` : Une puissante bibliothèque pour la manipulation et l'analyse de données.
- `tkinter` : L'interface Python standard pour la boîte à outils graphique Tk.

Assurez-vous d'avoir installé ces bibliothèques avant d'exécuter le script.


