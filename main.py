import pandas as pd
import win32com.client
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime, timedelta

def trouver_colonne_email(df):
    for col in df.columns:
        if any(df[col].astype(str).str.contains(r"@.*\.", na=False)):
            return col
    raise ValueError("Aucune colonne contenant des adresses e-mail n'a été trouvée.")

def analyser_contacts():
    filepath = filedialog.askopenfilename(title="Choisir le fichier CVS de contacts", filetypes=[("CSV files", "*.csv")])
    if not filepath:
        return

    try:
        years = int(simpledialog.askstring("Critère", "Aucun échange depuis combien d'années ?", initialvalue="2"))
        if years < 0:
            raise ValueError("Le nombre d'années doit être positif.")
    except (ValueError, TypeError):
        messagebox.showerror("Erreur", "Veuillez entrer un nombre d'années valide.")
        return

    try:
        contacts = pd.read_csv(filepath)
        contacts = contacts.dropna(axis=1, how='all')

        email_col = trouver_colonne_email(contacts)

        outlook = win32com.client.Dispatch("Outlook.Application")
        sent_folder = outlook.GetNamespace("MAPI").GetDefaultFolder(5)
        if not sent_folder.Items.Count:
            raise Exception("Le dossier des éléments envoyés est vide.")

        dernier_envoi = defaultdict(lambda: None)
        for item in sent_folder.Items:
            if hasattr(item, "To") and hasattr(item, "SentOn"):
                for dest in item.To.split(";"):
                    dest = dest.strip().lower()
                    if not dernier_envoi[dest] or item.SentOn > dernier_envoi[dest]:
                        dernier_envoi[dest] = item.SentOn

        contacts["dernier_envoi"] = contacts[email_col].map(lambda x: dernier_envoi.get(str(x).lower()))
        contacts["statut"] = contacts["dernier_envoi"].apply(
            lambda x: "Jamais contacté" if pd.isna(x) else "Contacté"
        )

        seuil_date = datetime.now() - timedelta(days=years * 365)
        contacts_filtrés = contacts[
            (contacts["dernier_envoi"].isna()) | 
            (contacts["dernier_envoi"] < seuil_date)
        ]

        output_path = filepath.replace(".csv", f"_sans_échange_{years}ans.csv")
        contacts_filtrés.to_csv(output_path, index=False, sep=';')

        messagebox.showinfo("Succès", f"Fichier généré avec {len(contacts_filtrés)} contacts sans échange depuis {years} ans :\n{output_path}")

    except Exception as e:
        print(f"Erreur rencontrée : {str(e)}")
        messagebox.showerror("Erreur", f"{str(e)}\nVérifiez qu'Outlook est installé, configuré et connecté.")

root = tk.Tk()
root.title("Analyse des contacts Outlook")
root.geometry("400x150")

label = tk.Label(root, text="Cliquez sur le bouton pour analyser les mails envoyés.")
label.pack(pady=20)

btn = tk.Button(root, text="Sélectionner le fichier CSV", command=analyser_contacts)
btn.pack()

root.mainloop()