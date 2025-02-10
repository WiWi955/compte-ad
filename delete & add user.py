import csv
import random
import string
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import logging
from datetime import datetime
from PIL import Image, ImageTk

# Configuration du fichier CSV
input_file = 'comptes-protec.csv'

# Configuration du logging
logging.basicConfig(filename='user_management.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def generate_password():
    """Génère un mot de passe aléatoire."""
    length = 13 - len("Protec")
    characters = string.ascii_letters + string.digits
    random_part = ''.join(random.choice(characters) for _ in range(length))
    symbol = random.choice(string.punctuation)
    return "Protec" + random_part + symbol

def add_user(nom, prenom, service):
    """Ajoute un utilisateur au fichier CSV."""
    try:
        with open(input_file, mode='r', newline='') as infile:
            reader = csv.DictReader(infile, delimiter=';')
            fieldnames = reader.fieldnames
            rows = list(reader)
       
        new_user = {
            'NOM': nom,
            'PRENOM': prenom,
            'IDENTIFIANT': f"{prenom.lower()}.{nom.lower()}",
            'PASSWORD': generate_password(),
            'service': service,
            'MESSAGERIE': f"{prenom[0].upper()}.{nom.upper()}@protec-groupe.com",
            'ExpirationDate': '02/07/2030',
            'ChangePassword': 'NON',
            'Script': f"{service}.bat"
        }
       
        rows.append(new_user)
       
        with open(input_file, mode='w', newline='') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            writer.writerows(rows)
       
        messagebox.showinfo("Succès", f"L'utilisateur {prenom} {nom} a été ajouté au fichier.")
        logging.info(f"Utilisateur ajouté : {prenom} {nom}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'ajout de l'utilisateur : {str(e)}")
        logging.error(f"Erreur lors de l'ajout de l'utilisateur : {str(e)}")

def delete_user(nom, prenom):
    """Supprime un utilisateur du fichier CSV."""
    try:
        with open(input_file, mode='r', newline='') as infile:
            reader = csv.DictReader(infile, delimiter=';')
            fieldnames = reader.fieldnames
            rows = [row for row in reader if row['NOM'] != nom or row['PRENOM'] != prenom]
       
        with open(input_file, mode='w', newline='') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            writer.writerows(rows)
       
        messagebox.showinfo("Succès", f"L'utilisateur {prenom} {nom} a été supprimé du fichier.")
        logging.info(f"Utilisateur supprimé : {prenom} {nom}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la suppression de l'utilisateur : {str(e)}")
        logging.error(f"Erreur lors de la suppression de l'utilisateur : {str(e)}")

def modifier_mots_de_passe():
    """Modifie tous les mots de passe des utilisateurs."""
    reponse = messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir modifier tous les mots de passe ?")
    if reponse:
        try:
            with open(input_file, mode='r', newline='', encoding='utf-8') as infile:
                reader = csv.DictReader(infile, delimiter=';')
                rows = list(reader)
                fieldnames = reader.fieldnames
           
            with open(input_file, mode='w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                writer.writeheader()
                for row in rows:
                    row['PASSWORD'] = generate_password()
                    writer.writerow(row)
           
            messagebox.showinfo("Succès", "Les mots de passe ont été mis à jour avec succès.")
            logging.info("Mots de passe de tous les utilisateurs modifiés.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la modification des mots de passe : {str(e)}")
            logging.error(f"Erreur lors de la modification des mots de passe : {str(e)}")

def modifier_mot_de_passe_user(nom, prenom):
    """Modifie le mot de passe d'un utilisateur spécifique."""
    try:
        with open(input_file, mode='r', newline='', encoding='utf-8') as infile:
            reader = csv.DictReader(infile, delimiter=';')
            rows = list(reader)
            fieldnames = reader.fieldnames
       
        user_found = False
        for row in rows:
            if row['NOM'] == nom and row['PRENOM'] == prenom:
                row['PASSWORD'] = generate_password()
                user_found = True
                break
       
        if user_found:
            with open(input_file, mode='w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
           
            messagebox.showinfo("Succès", f"Le mot de passe de {prenom} {nom} a été modifié avec succès.")
            logging.info(f"Mot de passe modifié pour l'utilisateur : {prenom} {nom}")
        else:
            messagebox.showerror("Erreur", f"L'utilisateur {prenom} {nom} n'a pas été trouvé.")
            logging.warning(f"Utilisateur non trouvé : {prenom} {nom}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la modification du mot de passe : {str(e)}")
        logging.error(f"Erreur lors de la modification du mot de passe : {str(e)}")

def synchroniser_ad():
    """Synchronise les utilisateurs avec Active Directory."""
    try:
        # Exécuter le script PowerShell de synchronisation
        result = subprocess.run(["powershell", "-File", "C:\\powershell\\synchroniser_ad.ps1"], capture_output=True, text=True)
       
        if result.returncode == 0:
            messagebox.showinfo("Succès", "Synchronisation avec AD réussie !")
            logging.info("Synchronisation avec AD réussie.")
        else:
            messagebox.showerror("Erreur", f"Erreur lors de la synchronisation : {result.stderr}")
            logging.error(f"Erreur lors de la synchronisation : {result.stderr}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Exception lors de la synchronisation : {str(e)}")
        logging.error(f"Exception lors de la synchronisation : {str(e)}")

def importer_utilisateurs():
    """Importe des utilisateurs à partir d'un fichier CSV."""
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        try:
            with open(file_path, mode='r', newline='', encoding='utf-8') as infile:
                reader = csv.DictReader(infile, delimiter=';')
                rows = list(reader)
           
            with open(input_file, mode='w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=reader.fieldnames, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
           
            messagebox.showinfo("Succès", "Utilisateurs importés avec succès.")
            logging.info("Utilisateurs importés depuis un fichier CSV.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'importation : {str(e)}")
            logging.error(f"Erreur lors de l'importation : {str(e)}")

def exporter_utilisateurs():
    """Exporte les utilisateurs vers un fichier CSV."""
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        try:
            with open(input_file, mode='r', newline='', encoding='utf-8') as infile:
                reader = csv.DictReader(infile, delimiter=';')
                rows = list(reader)
           
            with open(file_path, mode='w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=reader.fieldnames, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
           
            messagebox.showinfo("Succès", "Utilisateurs exportés avec succès.")
            logging.info("Utilisateurs exportés vers un fichier CSV.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'exportation : {str(e)}")
            logging.error(f"Erreur lors de l'exportation : {str(e)}")

def rechercher_utilisateur(nom, prenom):
    """Recherche un utilisateur dans le fichier CSV."""
    try:
        with open(input_file, mode='r', newline='', encoding='utf-8') as infile:
            reader = csv.DictReader(infile, delimiter=';')
            rows = [row for row in reader if row['NOM'] == nom and row['PRENOM'] == prenom]
       
        if rows:
            messagebox.showinfo("Résultat", f"Utilisateur trouvé : {rows[0]}")
            logging.info(f"Utilisateur trouvé : {rows[0]}")
        else:
            messagebox.showinfo("Résultat", "Aucun utilisateur trouvé.")
            logging.info("Aucun utilisateur trouvé.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la recherche : {str(e)}")
        logging.error(f"Erreur lors de la recherche : {str(e)}")

# Configuration de l'interface graphique
root = tk.Tk()
root.title("Gestion des Utilisateurs")
root.geometry('800x600')
root.resizable(False, False)

# Charger l'image en arrière-plan
background_image = Image.open('bg-main-dark-1440w.jpg')
bg = ImageTk.PhotoImage(background_image)

# Créer un label pour afficher l'image d'arrière-plan
background_label = tk.Label(root, image=bg)
background_label.place(relwidth=1, relheight=1)

# Conteneur pour les autres widgets
container = tk.Frame(root, bg='#f6f8fa')
container.place(relwidth=1, relheight=1)

# Le reste de ton interface (onglets et boutons) va ici
# Création du Notebook (onglets)
notebook = ttk.Notebook(container)
notebook.pack(expand=True, fill=tk.BOTH)

# Onglet pour ajouter/supprimer des utilisateurs
tab_ajouter_supprimer = ttk.Frame(notebook)
notebook.add(tab_ajouter_supprimer, text="Ajouter/Supprimer")

# Champs de saisie pour l'onglet Ajouter/Supprimer
label_nom = ttk.Label(tab_ajouter_supprimer, text="Nom :")
label_nom.grid(row=0, column=0, sticky="e", pady=10, padx=10)
entry_nom = ttk.Entry(tab_ajouter_supprimer, width=30)
entry_nom.grid(row=0, column=1, pady=10, padx=10, sticky="w")

label_prenom = ttk.Label(tab_ajouter_supprimer, text="Prénom :")
label_prenom.grid(row=1, column=0, sticky="e", pady=10, padx=10)
entry_prenom = ttk.Entry(tab_ajouter_supprimer, width=30)
entry_prenom.grid(row=1, column=1, pady=10, padx=10, sticky="w")

label_service = ttk.Label(tab_ajouter_supprimer, text="Service :")
label_service.grid(row=2, column=0, sticky="e", pady=10, padx=10)
services = [
    "Collaborateurs", "Commercial", "Comptabilité", "ContrôleNonDestructif", "DéveloppementDeSolutions", "Galva",
    "HygièneSécuritéEnvironnement", "Informatique", "MéthodesDeProduction", "Peinture", "Qualité",
    "RevueDeContrat", "R&D", "RessourcesHumaines", "Stagiaire", "TraitementDesMatériaux", "TraitementDeSurface",
    "WSI"
]
service_var = tk.StringVar()
dropdown_service = ttk.Combobox(tab_ajouter_supprimer, textvariable=service_var, values=services, state='readonly', width=28)
dropdown_service.grid(row=2, column=1, pady=10, padx=10, sticky="w")

# Boutons pour ajouter/supprimer
btn_ajouter = ttk.Button(tab_ajouter_supprimer, text="Ajouter", command=lambda: add_user(entry_nom.get(), entry_prenom.get(), service_var.get()))
btn_ajouter.grid(row=3, column=0, padx=10, pady=5)

btn_supprimer = ttk.Button(tab_ajouter_supprimer, text="Supprimer", command=lambda: delete_user(entry_nom.get(), entry_prenom.get()))
btn_supprimer.grid(row=3, column=1, padx=10, pady=5)

# Onglet pour modifier les mots de passe
tab_modifier_mdp = ttk.Frame(notebook)
notebook.add(tab_modifier_mdp, text="Modifier Mots de Passe")

# Champs de saisie pour modifier le mot de passe d'un utilisateur spécifique
label_nom_mdp = ttk.Label(tab_modifier_mdp, text="Nom :")
label_nom_mdp.grid(row=0, column=0, sticky="e", pady=10, padx=10)
entry_nom_mdp = ttk.Entry(tab_modifier_mdp, width=30)
entry_nom_mdp.grid(row=0, column=1, pady=10, padx=10, sticky="w")


label_prenom_mdp = ttk.Label(tab_modifier_mdp, text="Prénom :")
label_prenom_mdp.grid(row=1, column=0, sticky="e", pady=10, padx=10)
entry_prenom_mdp = ttk.Entry(tab_modifier_mdp, width=30)
entry_prenom_mdp.grid(row=1, column=1, pady=10, padx=10, sticky="w")

# Boutons pour modifier les mots de passe
btn_modifier_mdp_user = ttk.Button(tab_modifier_mdp, text="Modifier Mot de Passe", command=lambda: modifier_mot_de_passe_user(entry_nom_mdp.get(), entry_prenom_mdp.get()))
btn_modifier_mdp_user.grid(row=2, column=0, columnspan=2, pady=20)

btn_modifier_mdp_all = ttk.Button(tab_modifier_mdp, text="Modifier Tous les Mots de Passe", command=modifier_mots_de_passe)
btn_modifier_mdp_all.grid(row=3, column=0, columnspan=2, pady=20)

# Onglet pour synchroniser avec AD
tab_synchroniser = ttk.Frame(notebook)
notebook.add(tab_synchroniser, text="Synchroniser avec AD")

# Bouton pour synchroniser avec AD
btn_synchroniser = ttk.Button(tab_synchroniser, text="Synchroniser avec AD", command=synchroniser_ad)
btn_synchroniser.pack(pady=20)

# Onglet pour importer/exporter des utilisateurs
tab_importer_exporter = ttk.Frame(notebook)
notebook.add(tab_importer_exporter, text="Importer/Exporter")

# Boutons pour importer/exporter
btn_importer = ttk.Button(tab_importer_exporter, text="Importer Utilisateurs", command=importer_utilisateurs)
btn_importer.pack(pady=10)

btn_exporter = ttk.Button(tab_importer_exporter, text="Exporter Utilisateurs", command=exporter_utilisateurs)
btn_exporter.pack(pady=10)

# Onglet pour rechercher des utilisateurs
tab_rechercher = ttk.Frame(notebook)
notebook.add(tab_rechercher, text="Rechercher Utilisateur")

# Champs de saisie pour rechercher un utilisateur
label_nom_recherche = ttk.Label(tab_rechercher, text="Nom :")
label_nom_recherche.grid(row=0, column=0, sticky="e", pady=10, padx=10)
entry_nom_recherche = ttk.Entry(tab_rechercher, width=30)
entry_nom_recherche.grid(row=0, column=1, pady=10, padx=10, sticky="w")

label_prenom_recherche = ttk.Label(tab_rechercher, text="Prénom :")
label_prenom_recherche.grid(row=1, column=0, sticky="e", pady=10, padx=10)
entry_prenom_recherche = ttk.Entry(tab_rechercher, width=30)
entry_prenom_recherche.grid(row=1, column=1, pady=10, padx=10, sticky="w")

# Bouton pour rechercher un utilisateur
btn_rechercher = ttk.Button(tab_rechercher, text="Rechercher", command=lambda: rechercher_utilisateur(entry_nom_recherche.get(), entry_prenom_recherche.get()))
btn_rechercher.grid(row=2, column=0, columnspan=2, pady=20)

# Onglet pour gérer les groupes AD
tab_groupes_ad = ttk.Frame(notebook)
notebook.add(tab_groupes_ad, text="Gérer Groupes AD")

# Bouton pour gérer les groupes AD
btn_gerer_groupes_ad = ttk.Button(tab_groupes_ad, text="Gérer Groupes AD", command=gerer_groupes_ad)
btn_gerer_groupes_ad.pack(pady=20)

# Onglet pour gérer les permissions NTFS
tab_permissions_ntfs = ttk.Frame(notebook)
notebook.add(tab_permissions_ntfs, text="Gérer Permissions NTFS")

# Bouton pour gérer les permissions NTFS
btn_gerer_permissions_ntfs = ttk.Button(tab_permissions_ntfs, text="Gérer Permissions NTFS", command=gerer_permissions_ntfs)
btn_gerer_permissions_ntfs.pack(pady=20)

# Onglet pour exporter des rapports
tab_rapports = ttk.Frame(notebook)
notebook.add(tab_rapports, text="Exporter Rapports")

# Bouton pour exporter un rapport
btn_exporter_rapport = ttk.Button(tab_rapports, text="Exporter Rapport", command=exporter_rapport)
btn_exporter_rapport.pack(pady=20)

# Lancer l'interface
root.mainloop()

