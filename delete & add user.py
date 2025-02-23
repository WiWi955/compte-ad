import os
import csv
import random
import string
import subprocess
import logging
from datetime import datetime
import time
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel

# Configuration de l'encodage CSV pour une meilleure compatibilité Windows
CSV_ENCODING = 'utf-8-sig'
# Ajout du champ LOCKED pour la sécurité
INPUT_FILE = 'comptes-protec.csv'
DEFAULT_FIELDNAMES = ['NOM', 'PRENOM', 'IDENTIFIANT', 'PASSWORD', 'service',
                        'MESSAGERIE', 'ExpirationDate', 'ChangePassword', 'Script', 'LOCKED']

def check_csv_exists():
    """Vérifie l'existence du fichier CSV et le crée si nécessaire"""
    if not os.path.exists(INPUT_FILE):
        with open(INPUT_FILE, 'w', newline='', encoding=CSV_ENCODING) as f:
            writer = csv.DictWriter(f, fieldnames=DEFAULT_FIELDNAMES, delimiter=';')
            writer.writeheader()

def user_exists(nom, prenom):
    """Vérifie si l'utilisateur existe déjà"""
    try:
        with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
            reader = csv.DictReader(infile, delimiter=';')
            for row in reader:
                if row['NOM'].lower() == nom.lower() and row['PRENOM'].lower() == prenom.lower():
                    return True
    except FileNotFoundError:
        return False
    return False

def generate_password():
    """Génère un mot de passe conforme aux politiques de sécurité"""
    length = 14 - len("Protec")
    safe_symbols = '!#$%&*+-=?@^_'
    characters = string.ascii_uppercase + string.ascii_lowercase + string.digits + safe_symbols
    while True:
        password = "Protec" + ''.join(random.choice(characters) for _ in range(length-1))
        if (any(c.isupper() for c in password) and any(c.islower() for c in password) and
            any(c.isdigit() for c in password) and any(c in safe_symbols for c in password)):
            return password

def send_email_alert(subject, message):
    """Fonction simulée d'envoi d'alerte par email (ici, on logge l'alerte)"""
    logging.error(f"EMAIL ALERT: {subject} - {message}")

# Configuration du style pour le Tooltip avec ttk
root = tk.Tk()
style = ttk.Style(root)
style.configure("Tooltip.TLabel", background="#FFFFE0", foreground="black", font=("Arial", 9))
root.withdraw()  # Masque la fenêtre root utilisée pour le style

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip = Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tooltip, text=self.text, style="Tooltip.TLabel", relief='solid', borderwidth=1)
        label.pack()

    def hide(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestion des Utilisateurs - PROTEC")
        self.geometry("1200x800")
        self.minsize(1000, 700)

        self.colors = {
            "bg": "#1F1F2E",
            "sidebar": "#292B38",
            "accent": "#4EC9B0",
            "accent2": "#6B7A90",
            "text": "#FFFFFF",
            "expired": "#FFCCCC",   # Fond pour mot de passe expiré
            "locked": "#FFDD99"     # Fond pour compte verrouillé
        }
        self.configure(bg=self.colors["bg"])

        # Dictionnaire pour suivre les tentatives infructueuses de modification
        self.failed_attempts = {}

        # Barre latérale de navigation
        self.sidebar_frame = ctk.CTkFrame(self, width=250, fg_color=self.colors["sidebar"])
        self.sidebar_frame.pack(side="left", fill="y")

        # Zone de contenu principale
        self.content_frame = ctk.CTkFrame(self, fg_color=self.colors["bg"])
        self.content_frame.pack(side="right", fill="both", expand=True)

        self.create_sidebar()
        self.frames = {}
        self.create_frames()
        self.show_frame("Ajouter/Supprimer")

    def create_sidebar(self):
        title_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="PROTEC",
            font=("Helvetica", 24, "bold"),
            text_color=self.colors["accent"],
            bg_color=self.colors["sidebar"]
        )
        title_label.pack(pady=20)

        nav_items = ["Ajouter/Supprimer", "Mots de passe", "Synchronisation", "Import/Export", "Liste des Utilisateurs"]
        for item in nav_items:
            btn = ctk.CTkButton(
                self.sidebar_frame,
                text=item,
                command=lambda i=item: self.show_frame(i),
                fg_color=self.colors["accent"],
                hover_color=self.colors["accent2"],
                font=("Helvetica", 16),
                width=200
            )
            btn.pack(pady=10, padx=20)
            if item == "Synchronisation":
                temp_btn = ttk.Button(self.sidebar_frame, text="Synchronisation")
                Tooltip(temp_btn, "Lance la synchronisation avec Active Directory via PowerShell")

    def create_frames(self):
        # Onglet Ajouter/Supprimer
        frame_add_remove = ctk.CTkFrame(self.content_frame, fg_color=self.colors["bg"])
        self.frames["Ajouter/Supprimer"] = frame_add_remove
        self.create_add_remove_content(frame_add_remove)

        # Onglet Mots de passe
        frame_password = ctk.CTkFrame(self.content_frame, fg_color=self.colors["bg"])
        self.frames["Mots de passe"] = frame_password
        self.create_password_content(frame_password)

        # Onglet Synchronisation
        frame_sync = ctk.CTkFrame(self.content_frame, fg_color=self.colors["bg"])
        self.frames["Synchronisation"] = frame_sync
        self.create_sync_content(frame_sync)

        # Onglet Import/Export
        frame_import_export = ctk.CTkFrame(self.content_frame, fg_color=self.colors["bg"])
        self.frames["Import/Export"] = frame_import_export
        self.create_import_export_content(frame_import_export)

        # Onglet Liste des Utilisateurs
        frame_liste = ctk.CTkFrame(self.content_frame, fg_color=self.colors["bg"])
        self.frames["Liste des Utilisateurs"] = frame_liste
        self.create_liste_utilisateurs_content(frame_liste)

    def show_frame(self, frame_name):
        for frame in self.frames.values():
            frame.pack_forget()
        self.frames[frame_name].pack(fill="both", expand=True, padx=20, pady=20)

    def create_input_field(self, parent, label_text):
        frame = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        frame.pack(fill="x", pady=10)
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            text_color=self.colors["text"],
            font=("Helvetica", 16)
        )
        label.pack(side="left", padx=10)
        entry = ctk.CTkEntry(frame, font=("Helvetica", 16))
        entry.pack(side="right", fill="x", expand=True, padx=10)
        return entry

    # ------------------ Onglet: Ajouter/Supprimer ------------------
    def create_add_remove_content(self, parent):
        title = ctk.CTkLabel(
            parent,
            text="Ajouter / Supprimer Utilisateur",
            font=("Helvetica", 20, "bold"),
            text_color=self.colors["accent"]
        )
        title.pack(pady=10)

        self.entry_nom = self.create_input_field(parent, "Nom :")
        self.entry_prenom = self.create_input_field(parent, "Prénom :")

        services = [
            "Collaborateurs", "Commercial", "Comptabilité",
            "ContrôleNonDestructif", "DéveloppementDeSolutions",
            "Galva", "HygièneSécuritéEnvironnement", "Informatique",
            "MéthodesDeProduction", "Peinture", "Qualité",
            "RevueDeContrat", "R&D", "RessourcesHumaines",
            "Stagiaire", "TraitementDesMatériaux",
            "TraitementDeSurface", "WSI"
        ]
        frame_service = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        frame_service.pack(fill="x", pady=10)
        label_service = ctk.CTkLabel(
            frame_service,
            text="Service :",
            text_color=self.colors["text"],
            font=("Helvetica", 16)
        )
        label_service.pack(side="left", padx=10)
        self.service_combobox = ctk.CTkComboBox(
            frame_service,
            values=services,
            font=("Helvetica", 16),
            dropdown_font=("Helvetica", 16),
            button_color=self.colors["accent"],
            fg_color=self.colors["bg"],
            border_color=self.colors["accent2"]
        )
        self.service_combobox.pack(side="right", fill="x", expand=True, padx=10)

        frame_buttons = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        frame_buttons.pack(pady=20)
        btn_add = ctk.CTkButton(
            frame_buttons,
            text="Ajouter Utilisateur",
            command=self.add_user,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16)
        )
        btn_add.pack(side="left", padx=10)
        btn_delete = ctk.CTkButton(
            frame_buttons,
            text="Supprimer Utilisateur",
            command=self.delete_user,
            fg_color="#D35F5F",
            hover_color="#B74D4D",
            font=("Helvetica", 16)
        )
        btn_delete.pack(side="left", padx=10)

    def add_user(self):
        nom = self.entry_nom.get().strip()
        prenom = self.entry_prenom.get().strip()
        service = self.service_combobox.get().strip()
        if not nom or not prenom or not service:
            messagebox.showerror("Erreur", "Veuillez remplir tous les champs.")
            return
        if user_exists(nom, prenom):
            messagebox.showerror("Erreur", "Cet utilisateur existe déjà.")
            return
        try:
            with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
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
                'Script': f"{service}.bat",
                'LOCKED': "NON"
            }
            rows.append(new_user)
            with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
            messagebox.showinfo("Succès", f"L'utilisateur {prenom} {nom} a été ajouté.")
            logging.info(f"Utilisateur ajouté : {prenom} {nom}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'ajout de l'utilisateur : {str(e)}")
            logging.error(f"Erreur lors de l'ajout de l'utilisateur : {str(e)}")

    def delete_user(self):
        nom = self.entry_nom.get().strip()
        prenom = self.entry_prenom.get().strip()
        if not nom or not prenom:
            messagebox.showerror("Erreur", "Veuillez indiquer le nom et le prénom.")
            return
        try:
            with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                reader = csv.DictReader(infile, delimiter=';')
                fieldnames = reader.fieldnames
                rows = [row for row in reader if row['NOM'] != nom or row['PRENOM'] != prenom]
            with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
            messagebox.showinfo("Succès", f"L'utilisateur {prenom} {nom} a été supprimé.")
            logging.info(f"Utilisateur supprimé : {prenom} {nom}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la suppression de l'utilisateur : {str(e)}")
            logging.error(f"Erreur lors de la suppression de l'utilisateur : {str(e)}")

    # ------------------ Onglet: Mots de passe ------------------
    def create_password_content(self, parent):
        title = ctk.CTkLabel(
            parent,
            text="Gestion des Mots de Passe",
            font=("Helvetica", 20, "bold"),
            text_color=self.colors["accent"]
        )
        title.pack(pady=10)
        self.entry_nom_mdp = self.create_input_field(parent, "Nom :")
        self.entry_prenom_mdp = self.create_input_field(parent, "Prénom :")
        frame_buttons = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        frame_buttons.pack(pady=20)
        btn_modif = ctk.CTkButton(
            frame_buttons,
            text="Modifier mot de passe",
            command=self.modify_single_password,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16)
        )
        btn_modif.pack(side="left", padx=10)
        btn_modif_all = ctk.CTkButton(
            frame_buttons,
            text="Réinitialiser tous les mots de passe",
            command=self.modify_all_passwords,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16)
        )
        btn_modif_all.pack(side="left", padx=10)

    def modify_single_password(self):
        nom = self.entry_nom_mdp.get().strip()
        prenom = self.entry_prenom_mdp.get().strip()
        if not nom or not prenom:
            messagebox.showerror("Erreur", "Veuillez renseigner nom et prénom.")
            return
        try:
            with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                reader = csv.DictReader(infile, delimiter=';')
                rows = list(reader)
                fieldnames = reader.fieldnames
            user_found = False
            for row in rows:
                if row['NOM'] == nom and row['PRENOM'] == prenom:
                    if row.get("LOCKED", "NON") == "OUI":
                        messagebox.showerror("Compte verrouillé", "Ce compte est verrouillé.")
                        return
                    row['PASSWORD'] = generate_password()
                    user_found = True
                    self.failed_attempts[(nom, prenom)] = 0
                    break
            if user_found:
                with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
                messagebox.showinfo("Succès", f"Le mot de passe de {prenom} {nom} a été modifié.")
                logging.info(f"Mot de passe modifié pour : {prenom} {nom}")
            else:
                key = (nom, prenom)
                self.failed_attempts[key] = self.failed_attempts.get(key, 0) + 1
                if self.failed_attempts[key] >= 3:
                    for row in rows:
                        if row['NOM'] == nom and row['PRENOM'] == prenom:
                            row["LOCKED"] = "OUI"
                    with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                        writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                        writer.writeheader()
                        writer.writerows(rows)
                    send_email_alert("Compte verrouillé", f"Le compte de {prenom} {nom} a été verrouillé après plusieurs tentatives infructueuses.")
                    messagebox.showerror("Compte verrouillé", "Compte verrouillé après plusieurs tentatives infructueuses.")
                    logging.warning(f"Compte verrouillé pour : {prenom} {nom}")
                else:
                    messagebox.showerror("Erreur", f"L'utilisateur {prenom} {nom} n'a pas été trouvé.")
                    logging.warning(f"Utilisateur non trouvé : {prenom} {nom}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la modification du mot de passe : {str(e)}")
            logging.error(f"Erreur lors de la modification du mot de passe : {str(e)}")

    def modify_all_passwords(self):
        reponse = messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir modifier tous les mots de passe ?")
        if reponse:
            try:
                with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                    reader = csv.DictReader(infile, delimiter=';')
                    rows = list(reader)
                    fieldnames = reader.fieldnames
                with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                    writer.writeheader()
                    for row in rows:
                        if row.get("LOCKED", "NON") == "NON":
                            row['PASSWORD'] = generate_password()
                        writer.writerow(row)
                messagebox.showinfo("Succès", "Tous les mots de passe ont été modifiés.")
                logging.info("Tous les mots de passe ont été modifiés.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la modification globale : {str(e)}")
                logging.error(f"Erreur lors de la modification globale : {str(e)}")

    # ------------------ Onglet: Synchronisation ------------------
    def create_sync_content(self, parent):
        title = ctk.CTkLabel(
            parent,
            text="Synchronisation Active Directory",
            font=("Helvetica", 20, "bold"),
            text_color=self.colors["accent"]
        )
        title.pack(pady=10)
        btn_sync = ctk.CTkButton(
            parent,
            text="LANCER LA SYNCHRONISATION AD",
            command=self.sync_ad,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 18),
            width=300,
            height=50
        )
        btn_sync.pack(pady=40)

    def sync_ad(self):
        progress_win = Toplevel(self)
        progress_win.title("Synchronisation en cours...")
        progress_bar = ctk.CTkProgressBar(progress_win, width=300)
        progress_bar.pack(padx=20, pady=20)
        progress_bar.set(0)
        self.update()
        for i in range(1, 11):
            progress_bar.set(i/10)
            progress_win.update()
            time.sleep(0.3)
        progress_win.destroy()
        try:
            subprocess.Popen(["cmd.exe", "/c", "start", "powershell", "-NoExit", "-File", "C:\\powershell\\synchroniser_ad.ps1"], shell=True)
            messagebox.showinfo("Succès", "Synchronisation AD lancée.")
            logging.info("Synchronisation AD lancée.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la synchronisation : {str(e)}")
            logging.error(f"Erreur lors de la synchronisation : {str(e)}")

    # ------------------ Onglet: Import/Export ------------------
    def create_import_export_content(self, parent):
        title = ctk.CTkLabel(
            parent,
            text="Import / Export Utilisateurs",
            font=("Helvetica", 20, "bold"),
            text_color=self.colors["accent"]
        )
        title.pack(pady=10)
        frame_buttons = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        frame_buttons.pack(pady=20)
        btn_import = ctk.CTkButton(
            frame_buttons,
            text="Importer Utilisateurs",
            command=self.import_users,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16),
            width=200,
            height=40
        )
        btn_import.pack(side="left", padx=10)
        btn_export = ctk.CTkButton(
            frame_buttons,
            text="Exporter Utilisateurs",
            command=self.export_users,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16),
            width=200,
            height=40
        )
        btn_export.pack(side="left", padx=10)
        btn_report = ctk.CTkButton(
            frame_buttons,
            text="Exporter rapport",
            command=self.export_report,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent2"],
            font=("Helvetica", 16),
            width=200,
            height=40
        )
        btn_report.pack(side="left", padx=10)

    def import_users(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            progress_win = Toplevel(self)
            progress_win.title("Import en cours...")
            progress_bar = ctk.CTkProgressBar(progress_win, width=300)
            progress_bar.pack(padx=20, pady=20)
            progress_bar.set(0)
            self.update()
            for i in range(1, 11):
                progress_bar.set(i/10)
                progress_win.update()
                time.sleep(0.2)
            progress_win.destroy()
            try:
                with open(file_path, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                    reader = csv.DictReader(infile, delimiter=';')
                    rows = list(reader)
                with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=reader.fieldnames, delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
                messagebox.showinfo("Succès", "Utilisateurs importés avec succès.")
                logging.info("Utilisateurs importés depuis un fichier CSV.")
                self.load_user_table()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'importation : {str(e)}")
                logging.error(f"Erreur lors de l'importation : {str(e)}")

    def export_users(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                    reader = csv.DictReader(infile, delimiter=';')
                    rows = list(reader)
                with open(file_path, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=reader.fieldnames, delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
                messagebox.showinfo("Succès", "Utilisateurs exportés avec succès.")
                logging.info("Utilisateurs exportés vers un fichier CSV.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'exportation : {str(e)}")
                logging.error(f"Erreur lors de l'exportation : {str(e)}")

    def export_report(self):
        """Exporte un rapport détaillé (ici, un CSV) sur l'état des comptes utilisateurs."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", title="Enregistrer le rapport", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                    reader = csv.DictReader(infile, delimiter=';')
                    rows = list(reader)
                with open(file_path, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=DEFAULT_FIELDNAMES, delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
                messagebox.showinfo("Succès", "Rapport exporté avec succès.")
                logging.info("Rapport exporté.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'exportation du rapport : {str(e)}")
                logging.error(f"Erreur lors de l'exportation du rapport : {str(e)}")

    # ------------------ Onglet: Liste des Utilisateurs ------------------
    def create_liste_utilisateurs_content(self, parent):
        """Crée et affiche un tableau moderne des utilisateurs avec en-têtes fixes et filtres de recherche."""
        title = ctk.CTkLabel(
            parent,
            text="Liste des Utilisateurs",
            font=("Helvetica", 20, "bold"),
            text_color=self.colors["accent"]
        )
        title.pack(pady=10)
        
        # Cadre des filtres de recherche
        filter_frame = ctk.CTkFrame(parent, fg_color=self.colors["bg"])
        filter_frame.pack(fill="x", padx=20, pady=(0,10))
        
        self.search_entry = ctk.CTkEntry(
            filter_frame,
            font=("Helvetica", 16),
            placeholder_text="Rechercher (Nom, Prénom, Mdp, Email)"
        )
        self.search_entry.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda event: self.load_user_table())
        
        # Liste des services mise à jour
        services = ["Tous"] + [
            "Collaborateurs", "Commercial", "Comptabilité",
            "ContrôleNonDestructif", "DéveloppementDeSolutions", "Galva",
            "HygièneSécuritéEnvironnement", "Informatique", "MéthodesDeProduction",
            "Peinture", "Qualité", "RevueDeContrat", "R&D", "RessourcesHumaines",
            "Stagiaire", "TraitementDesMatériaux", "TraitementDeSurface", "WSI"
        ]
        self.filter_service = ctk.CTkComboBox(
            filter_frame,
            values=services,
            font=("Helvetica", 16),
            width=200
        )
        self.filter_service.set("Tous")
        self.filter_service.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.filter_service.bind("<<ComboboxSelected>>", lambda event: self.load_user_table())
        
        filter_frame.grid_columnconfigure(0, weight=1)
        filter_frame.grid_columnconfigure(1, weight=1)
        
        # En-têtes du tableau
        header_frame = ctk.CTkFrame(parent, fg_color=self.colors["sidebar"], corner_radius=8)
        header_frame.pack(fill="x", padx=20)
        headings = ["Nom", "Prénom", "Mot de passe", "Email", "Actions"]
        for j, heading in enumerate(headings):
            header = ctk.CTkLabel(
                header_frame,
                text=heading,
                font=("Helvetica", 16, "bold"),
                text_color=self.colors["accent"],
                fg_color=self.colors["sidebar"],
                anchor="center"
            )
            header.grid(row=0, column=j, padx=10, pady=10, sticky="nsew")
            header_frame.grid_columnconfigure(j, weight=1)
        
        # Zone scrollable pour les données
        self.scroll_frame = ctk.CTkScrollableFrame(
            parent,
            fg_color=self.colors["bg"],
            corner_radius=8,
            border_width=1,
            border_color=self.colors["accent2"]
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0,20))
        for j in range(len(headings)):
            self.scroll_frame.grid_columnconfigure(j, weight=1)
        
        self.load_user_table()

    def load_user_table(self):
        """Charge et affiche les utilisateurs selon les filtres de recherche."""
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        search_term = self.search_entry.get().strip().lower() if self.search_entry.get() else None
        filter_service = self.filter_service.get().strip().lower() if self.filter_service.get() != "Tous" else None
        
        try:
            with open(INPUT_FILE, 'r', newline='', encoding=CSV_ENCODING) as f:
                reader = csv.DictReader(f, delimiter=';')
                row_index = 0
                for row in reader:
                    # Filtre recherche
                    if search_term:
                        if (search_term not in row['NOM'].lower() and
                            search_term not in row['PRENOM'].lower() and
                            search_term not in row['PASSWORD'].lower() and
                            search_term not in row['MESSAGERIE'].lower()):
                            continue
                    
                    # Filtre service (égalité exacte)
                    if filter_service:
                        if row['service'].lower() != filter_service:
                            continue
                    
                    # Couleur de fond selon état
                    bg_color = self.colors["bg"]
                    try:
                        exp_date = datetime.strptime(row['ExpirationDate'], "%d/%m/%Y")
                        if exp_date < datetime.today():
                            bg_color = self.colors["expired"]
                    except Exception:
                        pass
                    if row.get("LOCKED", "NON") == "OUI":
                        bg_color = self.colors["locked"]
                    
                    # Affichage des données
                    label_nom = ctk.CTkLabel(
                        self.scroll_frame,
                        text=row['NOM'],
                        font=("Helvetica", 14),
                        text_color=self.colors["text"],
                        fg_color=bg_color,
                        anchor="center"
                    )
                    label_nom.grid(row=row_index, column=0, padx=10, pady=5, sticky="nsew")
                    
                    label_prenom = ctk.CTkLabel(
                        self.scroll_frame,
                        text=row['PRENOM'],
                        font=("Helvetica", 14),
                        text_color=self.colors["text"],
                        fg_color=bg_color,
                        anchor="center"
                    )
                    label_prenom.grid(row=row_index, column=1, padx=10, pady=5, sticky="nsew")
                    
                    label_password = ctk.CTkLabel(
                        self.scroll_frame,
                        text=row['PASSWORD'],
                        font=("Helvetica", 14),
                        text_color=self.colors["text"],
                        fg_color=bg_color,
                        anchor="center"
                    )
                    label_password.grid(row=row_index, column=2, padx=10, pady=5, sticky="nsew")
                    
                    label_email = ctk.CTkLabel(
                        self.scroll_frame,
                        text=row['MESSAGERIE'],
                        font=("Helvetica", 14),
                        text_color=self.colors["text"],
                        fg_color=bg_color,
                        anchor="center"
                    )
                    label_email.grid(row=row_index, column=3, padx=10, pady=5, sticky="nsew")
                    
                    btn_edit = ctk.CTkButton(
                        self.scroll_frame,
                        text="Editer",
                        font=("Helvetica", 14),
                        command=lambda r=row: self.edit_user(r),
                        width=80
                    )
                    btn_edit.grid(row=row_index, column=4, padx=10, pady=5, sticky="nsew")
                    
                    row_index += 1
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'affichage du tableau : {str(e)}")

    def edit_user(self, user_data):
        """Ouvre une fenêtre d'édition pour modifier les informations d'un utilisateur."""
        edit_win = Toplevel(self)
        edit_win.title("Modifier utilisateur")
        edit_win.geometry("400x400")
        
        entry_nom = ctk.CTkEntry(edit_win, font=("Helvetica", 16))
        entry_nom.pack(pady=10, padx=20, fill="x")
        entry_nom.insert(0, user_data['NOM'])
        
        entry_prenom = ctk.CTkEntry(edit_win, font=("Helvetica", 16))
        entry_prenom.pack(pady=10, padx=20, fill="x")
        entry_prenom.insert(0, user_data['PRENOM'])
        
        entry_service = ctk.CTkEntry(edit_win, font=("Helvetica", 16))
        entry_service.pack(pady=10, padx=20, fill="x")
        entry_service.insert(0, user_data['service'])
        
        entry_email = ctk.CTkEntry(edit_win, font=("Helvetica", 16))
        entry_email.pack(pady=10, padx=20, fill="x")
        entry_email.insert(0, user_data['MESSAGERIE'])
        
        entry_exp = ctk.CTkEntry(edit_win, font=("Helvetica", 16))
        entry_exp.pack(pady=10, padx=20, fill="x")
        entry_exp.insert(0, user_data['ExpirationDate'])
        
        def save_changes():
            new_nom = entry_nom.get().strip()
            new_prenom = entry_prenom.get().strip()
            new_service = entry_service.get().strip()
            new_email = entry_email.get().strip()
            new_exp = entry_exp.get().strip()
            if not (new_nom and new_prenom and new_service and new_email and new_exp):
                messagebox.showerror("Erreur", "Tous les champs sont obligatoires.")
                return
            try:
                with open(INPUT_FILE, mode='r', newline='', encoding=CSV_ENCODING) as infile:
                    reader = csv.DictReader(infile, delimiter=';')
                    rows = list(reader)
                    fieldnames = reader.fieldnames
                for row in rows:
                    if row['NOM'] == user_data['NOM'] and row['PRENOM'] == user_data['PRENOM']:
                        row['NOM'] = new_nom
                        row['PRENOM'] = new_prenom
                        row['service'] = new_service
                        row['MESSAGERIE'] = new_email
                        row['ExpirationDate'] = new_exp
                        break
                with open(INPUT_FILE, mode='w', newline='', encoding=CSV_ENCODING) as outfile:
                    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
                messagebox.showinfo("Succès", "Les modifications ont été sauvegardées.")
                logging.info(f"Utilisateur modifié : {new_prenom} {new_nom}")
                edit_win.destroy()
                self.load_user_table()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {str(e)}")
                logging.error(f"Erreur lors de la sauvegarde des modifications : {str(e)}")
                
        btn_save = ctk.CTkButton(edit_win, text="Sauvegarder", command=save_changes, font=("Helvetica", 16))
        btn_save.pack(pady=20)

if __name__ == "__main__":
    check_csv_exists()  # Vérifie et crée le fichier CSV si nécessaire
    app = App()
    app.mainloop()
