import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Importer les bibliothèques nécessaires

def create_excel_copies():
    year = int(year_entry.get())
    month = int(month_entry.get())
    # Spécifier le chemin vers votre fichier source
    source_file = 'C:/Users/takacsc/Documents/Projets/File-Multipli/source-file/testing.xlsx'

    # Obtenir le nombre de jours dans le mois et l'année spécifiés
    num_days = pd.Period(str(year) + '-' + str(month)).days_in_month

    # Générer une plage de dates pour le mois donné
    dates = pd.date_range(
        start=f'{year}-{month}-01', periods=num_days, freq='D')

    # Formater la date comme requis pour le nom de fichier
    formatted_dates = dates.strftime('%d %B %Y')

    # Créer le répertoire de destination
    month_name = pd.Timestamp(year=year, month=month, day=1).strftime('%B')
    directory_name = f"{month_name} {year}"
    os.makedirs(directory_name, exist_ok=True)

    # Parcourir les dates formatées et créer des copies dans le répertoire de destination
    for formatted_date in formatted_dates:
        file_name = f"Rapport du {formatted_date}.xlsx"
        destination_file = os.path.join(directory_name, file_name)
        shutil.copy(source_file, destination_file)
        status_label.config(text=f"Création de la copie : {destination_file}")


# Créer la fenêtre Tkinter
window = tk.Tk()
window.title("Copieur de fichiers Excel")

# Définir le titre de la fenêtre
title_label = tk.Label(window, text="Copieur de fichiers Excel",
                       font=("Helvetica", 16, "bold"))
title_label.pack()

# Définir les dimensions de la fenêtre
window.geometry("400x200")

# Centrer la fenêtre à l'écran
window_width = window.winfo_width()
window_height = window.winfo_height()
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
window.geometry(f"+{x}+{y}")

# Créer un cadre pour les champs d'entrée
input_frame = tk.Frame(window)
input_frame.pack(pady=10)

year_label = tk.Label(input_frame, text="Année:")
year_label.grid(row=0, column=0, padx=5, pady=5)
year_entry = tk.Entry(input_frame)
year_entry.grid(row=0, column=1, padx=5, pady=5)

month_label = tk.Label(input_frame, text="Mois:")
month_label.grid(row=1, column=0, padx=5, pady=5)
month_entry = tk.Entry(input_frame)
month_entry.grid(row=1, column=1, padx=5, pady=5)

# Créer le bouton pour démarrer le processus
process_button = tk.Button(window, text="Démarrer", command=create_excel_copies)
process_button.pack(pady=10)

# Créer une étiquette pour afficher l'état ou les messages
status_label = tk.Label(window, text="", fg="green")
status_label.pack()


# Exécuter la boucle d'événements Tkinter
window.mainloop()
