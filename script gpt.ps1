import csv
import random
import string
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import logging
from datetime import datetime

# Configuration du fichier CSV
input_file = 'comptes-protec.csv'

# Configuration du logging
logging.basicConfig(filename='user_management.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def generate_password():
    length = 13 - len("Protec")
    characters = string.ascii_letters + string.digits
    random_part = ''.join(random.choice(characters) for _ in range(length))
    symbol = random.choice(string.punctuation)
    return "Protec" + random_part + symbol

def add_user(nom, prenom, service):
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

# Configuration de l'interface graphique
root = tk.Tk()
root.title("Gestion des Utilisateurs")
root.geometry('800x600')
root.resizable(False, False)
root.configure(bg='#f6f8fa')

# Style personnalisé
style = ttk.Style()
style.theme_use('default')

# Création du Notebook (onglets)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill=tk.BOTH)

# Onglet pour ajouter/supprimer des utilisateurs
tab_ajouter_supprimer = ttk.Frame(notebook)
notebook.add(tab_ajouter_supprimer, text="Ajouter/Supprimer")

# Onglet pour modifier les mots de passe
tab_modifier_mdp = ttk.Frame(notebook)
notebook.add(tab_modifier_mdp, text="Modifier Mots de Passe")

# Onglet pour synchroniser avec AD
tab_synchroniser = ttk.Frame(notebook)
notebook.add(tab_synchroniser, text="Synchroniser avec AD")

# Onglet pour importer/exporter des utilisateurs
tab_importer_exporter = ttk.Frame(notebook)
notebook.add(tab_importer_exporter, text="Importer/Exporter")

# Ajouter les éléments dans chaque onglet (comme dans votre code original)
# ...

root.mainloop()