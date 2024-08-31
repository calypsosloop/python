import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage

# Charger les données Excel
file_path = 'Transducer Compatibility.xlsx'

# Lire la feuille spécifique contenant les données pertinentes
df = pd.read_excel(file_path, sheet_name='Feuille 1')  # Remplacez 'Feuille 1' par le nom correct de la feuille

# Afficher les noms des colonnes pour vérifier
print("Noms des colonnes dans la feuille :", df.columns)

# Nettoyer les données (supprimer les lignes avec NaN dans les colonnes importantes)
df = df.dropna(subset=['Brand', 'Sceen Model', 'Mount Type', 'Depth Range', 'Transducer model', 'Compatibility'])

# Remplacer les espaces par des underscores et uniformiser la casse dans les données
df['Brand'] = df['Brand'].str.replace(' ', '_').str.upper()
df['Sceen Model'] = df['Sceen Model'].str.replace(' ', '_').str.upper()
df['Mount Type'] = df['Mount Type'].str.replace(' ', '_').str.upper()
df['Depth Range'] = df['Depth Range'].str.replace(' ', '_').str.upper()
df['Transducer model'] = df['Transducer model'].str.replace(' ', '_').str.upper()
df['Compatibility'] = df['Compatibility'].str.replace(' ', '_').str.upper()

# Filtrer les transducteurs compatibles
transducers_compatibles = df[df['Compatibility'] == 'YES']['Transducer model'].unique()


# Fonction pour mettre à jour les modèles d'écran en fonction de la marque sélectionnée
def update_modeles(event):
    marque = marque_combobox.get().strip().upper()

    # Filtrer les modèles d'écran disponibles pour la marque sélectionnée
    modeles_filtrés = df[df['Brand'] == marque]['Sceen Model'].unique()

    # Mettre à jour les valeurs de la combobox des modèles d'écran
    modele_combobox['values'] = list(modeles_filtrés)
    modele_combobox.set('')  # Réinitialiser la sélection du modèle d'écran


# Fonction pour afficher les compatibilités par marque/modèle
def afficher_compatibilite():
    marque = marque_combobox.get().strip().upper()
    modele = modele_combobox.get().strip().upper()
    mount_type = mount_combobox.get().strip().upper()
    depth_range = depth_combobox.get().strip().upper()

    # Afficher les valeurs filtrées pour le débogage
    print(f"Marque sélectionnée: {marque}")
    print(f"Modèle sélectionné: {modele}")
    print(f"Type de montage sélectionné: {mount_type}")
    print(f"Plage de profondeur sélectionnée: {depth_range}")

    # Filtrer les données en fonction des choix de l'utilisateur
    resultats = df[(df['Brand'] == marque) & (df['Sceen Model'] == modele)]

    if mount_type:  # Appliquer le filtre seulement si mount_type n'est pas vide
        resultats = resultats[resultats['Mount Type'] == mount_type]
    if depth_range:  # Appliquer le filtre seulement si depth_range n'est pas vide
        resultats = resultats[resultats['Depth Range'] == depth_range]

    # Trier les résultats pour que les sondes compatibles soient en premier
    resultats = resultats.sort_values(by='Compatibility', ascending=False)

    # Effacer le contenu précédent de la Textbox
    resultat_text.delete(1.0, tk.END)

    # Afficher les sondes compatibles et non compatibles
    for index, row in resultats.iterrows():
        transducer_info = row['Transducer model']
        compatibilite = "Oui" if row['Compatibility'] == 'YES' else "Non"
        couleur = 'green' if compatibilite == "Oui" else 'black'

        # Vérifier si la note est présente et valide
        notes = row.get('Notes', '')
        if pd.notna(notes) and notes != 'Aucune':
            ligne = f"- Sonde: {transducer_info} | Compatible: {compatibilite} | Notes: {notes}\n"
        else:
            ligne = f"- Sonde: {transducer_info} | Compatible: {compatibilite}\n"

        resultat_text.insert(tk.END, ligne, couleur)

        # Ajouter un saut de ligne entre chaque entrée pour plus de lisibilité
        resultat_text.insert(tk.END, "\n")

    # Vérifier s'il n'y a pas de résultats
    if resultats.empty:
        resultat_text.insert(tk.END, "Aucune sonde compatible trouvée avec des notes spécifiques\n", 'red')


# Fonction pour afficher les compatibilités par transducteur
def afficher_par_transducteur():
    transducer = transducer_combobox.get().strip().upper()

    # Afficher les valeurs filtrées pour le débogage
    print(f"Transducteur sélectionné: {transducer}")

    # Filtrer les données en fonction du transducteur sélectionné
    resultats = df[(df['Transducer model'] == transducer) & (df['Compatibility'] == 'YES')]

    # Effacer le contenu précédent de la Textbox
    resultat_text2.delete(1.0, tk.END)

    # Vérifier si des résultats ont été trouvés
    if not resultats.empty:
        for index, row in resultats.iterrows():
            marque = row['Brand']
            modele = row['Sceen Model']
            compatibilite = "Oui" if row['Compatibility'] == 'YES' else "Non"
            couleur = 'green' if compatibilite == "Oui" else 'black'

            ligne = f"- Marque: {marque} | Modèle: {modele} | Compatible: {compatibilite}\n"
            resultat_text2.insert(tk.END, ligne, couleur)

            # Ajouter un saut de ligne entre chaque entrée pour plus de lisibilité
            resultat_text2.insert(tk.END, "\n")
    else:
        resultat_text2.insert(tk.END, "Aucune compatibilité trouvée pour ce transducteur\n", 'red')


# Fonction pour effacer les sélections et les résultats
def clear_selections():
    marque_combobox.set('')
    modele_combobox.set('')
    mount_combobox.set('')
    depth_combobox.set('')
    transducer_combobox.set('')
    resultat_text.delete(1.0, tk.END)
    resultat_text2.delete(1.0, tk.END)


# Fonction pour afficher la page de filtrage par marque/modèle
def show_page1():
    page1_frame.pack(fill='both', expand=True)
    page2_frame.pack_forget()


# Fonction pour afficher la page de filtrage par transducteur
def show_page2():
    page1_frame.pack_forget()
    page2_frame.pack(fill='both', expand=True)


# Créer la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Transducer Compatibility Tool")
fenetre.geometry("900x700")  # Agrandir la fenêtre pour plus d'espace
fenetre.configure(bg="#e0e0e0")  # Couleur de fond plus moderne

# Appliquer un thème moderne
style = ttk.Style()
style.theme_use('clam')  # Utiliser un thème moderne
style.configure('TLabel', font=('Arial', 10))
style.configure('TButton', font=('Arial', 10), background='#4CAF50', foreground='white')
style.configure('TCombobox', font=('Arial', 10))

# Ajouter une icône (optionnel, nécessite un fichier .ico valide)
# fenetre.iconbitmap('path/to/icon.ico')

# Créer les cadres pour les deux pages
page1_frame = tk.Frame(fenetre, bg="#e0e0e0")
page2_frame = tk.Frame(fenetre, bg="#e0e0e0")

# ------ Page 1: Filtrage par Marque/Modèle ------

# Cadre principal pour Page 1
main_frame1 = tk.Frame(page1_frame, bg="#e0e0e0")
main_frame1.pack(pady=20, padx=20)

# Combobox pour la sélection des marques
tk.Label(main_frame1, text="Select Brand:", bg="#e0e0e0", font=('Arial', 12)).grid(row=0, column=0, padx=10, pady=5,
                                                                                   sticky='w')
marque_combobox = ttk.Combobox(main_frame1, values=list(df['Brand'].unique()), width=30)
marque_combobox.grid(row=0, column=1, padx=10, pady=5)
marque_combobox.bind("<<ComboboxSelected>>",
                     update_modeles)  # Lier l'événement de sélection pour mettre à jour les modèles

# Combobox pour la sélection des modèles d'écran
tk.Label(main_frame1, text="Select Model:", bg="#e0e0e0", font=('Arial', 12)).grid(row=1, column=0, padx=10, pady=5,
                                                                                   sticky='w')
modele_combobox = ttk.Combobox(main_frame1, values=[], width=30)
modele_combobox.grid(row=1, column=1, padx=10, pady=5)

# Combobox pour la sélection du type de montage
tk.Label(main_frame1, text="Select Mount Type:", bg="#e0e0e0", font=('Arial', 12)).grid(row=2, column=0, padx=10,
                                                                                        pady=5, sticky='w')
mount_combobox = ttk.Combobox(main_frame1, values=list(df['Mount Type'].unique()), width=30)
mount_combobox.grid(row=2, column=1, padx=10, pady=5)

# Combobox pour la sélection de la plage de profondeur
tk.Label(main_frame1, text="Select Depth Range:", bg="#e0e0e0", font=('Arial', 12)).grid(row=3, column=0, padx=10,
                                                                                         pady=5, sticky='w')
depth_combobox = ttk.Combobox(main_frame1, values=list(df['Depth Range'].unique()), width=30)
depth_combobox.grid(row=3, column=1, padx=10, pady=5)

# Bouton pour afficher la compatibilité
bouton = ttk.Button(main_frame1, text="Show Compatibility", command=afficher_compatibilite)
bouton.grid(row=4, column=0, columnspan=2, pady=10)

# Bouton pour effacer les sélections et les résultats
clear_button1 = ttk.Button(main_frame1, text="Clear Selections", command=clear_selections)
clear_button1.grid(row=5, column=0, columnspan=2, pady=5)

# Bouton pour passer à la page de filtrage par transducteur
switch_button1 = ttk.Button(main_frame1, text="Go to Transducer Filter", command=show_page2)
switch_button1.grid(row=6, column=0, columnspan=2, pady=10)

# Textbox pour afficher les résultats avec Scrollbar pour Page 1
result_frame1 = tk.Frame(page1_frame, bg="#e0e0e0")
result_frame1.pack(fill='both', expand=True, pady=10, padx=20)

resultat_text = tk.Text(result_frame1, height=15, width=70, wrap=tk.WORD, font=('Arial', 10))
resultat_text.pack(side='left', fill='both', expand=True)

# Scrollbar pour la Textbox
scrollbar1 = ttk.Scrollbar(result_frame1, command=resultat_text.yview)
resultat_text['yscrollcommand'] = scrollbar1.set
scrollbar1.pack(side='right', fill='y')

# ------ Page 2: Filtrage par Transducteur ------

# Cadre principal pour Page 2
main_frame2 = tk.Frame(page2_frame, bg="#e0e0e0")
main_frame2.pack(pady=20, padx=20)

# Combobox pour la sélection des transducteurs compatibles
tk.Label(main_frame2, text="Select Transducer:", bg="#e0e0e0", font=('Arial', 12)).grid(row=0, column=0, padx=10,
                                                                                        pady=5, sticky='w')
transducer_combobox = ttk.Combobox(main_frame2, values=list(transducers_compatibles), width=30)
transducer_combobox.grid(row=0, column=1, padx=10, pady=5)

# Bouton pour afficher la compatibilité par transducteur
transducer_button = ttk.Button(main_frame2, text="Show by Transducer", command=afficher_par_transducteur)
transducer_button.grid(row=1, column=0, columnspan=2, pady=10)

# Bouton pour effacer les sélections et les résultats
clear_button2 = ttk.Button(main_frame2, text="Clear Selections", command=clear_selections)
clear_button2.grid(row=2, column=0, columnspan=2, pady=5)

# Bouton pour retourner à la page de filtrage par marque/modèle
switch_button2 = ttk.Button(main_frame2, text="Go to Brand/Model Filter", command=show_page1)
switch_button2.grid(row=3, column=0, columnspan=2, pady=10)

# Textbox pour afficher les résultats avec Scrollbar pour Page 2
result_frame2 = tk.Frame(page2_frame, bg="#e0e0e0")
result_frame2.pack(fill='both', expand=True, pady=10, padx=20)

resultat_text2 = tk.Text(result_frame2, height=15, width=70, wrap=tk.WORD, font=('Arial', 10))
resultat_text2.pack(side='left', fill='both', expand=True)

# Scrollbar pour la Textbox
scrollbar2 = ttk.Scrollbar(result_frame2, command=resultat_text2.yview)
resultat_text2['yscrollcommand'] = scrollbar2.set
scrollbar2.pack(side='right', fill='y')

# Configuration des tags de couleur et de style
resultat_text.tag_configure('green', foreground='green', font=('Arial', 10, 'bold'))
resultat_text.tag_configure('red', foreground='red', font=('Arial', 10, 'bold'))
resultat_text.tag_configure('black', foreground='black', font=('Arial', 10))
resultat_text2.tag_configure('green', foreground='green', font=('Arial', 10, 'bold'))
resultat_text2.tag_configure('red', foreground='red', font=('Arial', 10))
resultat_text2.tag_configure('black', foreground='black', font=('Arial', 10))

# Afficher la première page par défaut
show_page1()

# Lancer la boucle principale de l'interface
fenetre.mainloop()
