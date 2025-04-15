import pandas as pd
import win32com.client
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox

def analyser_contacts():
    filepath = filedialog.askopenfilename(title="Choisir le fichier CSV de contacts", filetypes=[("CSV files", "*.csv")])
    if not filepath:
        return

    try:
        contacts = pd.read_csv(filepath)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        sent_folder = outlook.GetDefaultFolder(5)

        dernier_envoi = defaultdict(lambda: None)

        for item in sent_folder.Items:
            if hasattr(item, "To") and hasattr(item, "SentOn"):
                for dest in item.To.split(";"):
                    dest = dest.strip().lower()
                    if not dernier_envoi[dest] or item.SentOn > dernier_envoi[dest]:
                        dernier_envoi[dest] = item.SentOn

        contacts["dernier_envoi"] = contacts["Adresse e-mail"].map(lambda x: dernier_envoi.get(str(x).lower()))
        output_path = filepath.replace(".csv", "_avec_dates.csv")
        contacts.to_csv(output_path, index=False)

        messagebox.showinfo("Succès", f"Fichier généré :\n{output_path}")

    except Exception as e:
        messagebox.showerror("Erreur", str(e))


# Interface graphique
root = tk.Tk()
root.title("Analyse des contacts Outlook")
root.geometry("400x150")

label = tk.Label(root, text="Cliquez sur le bouton pour analyser les mails envoyés.")
label.pack(pady=20)

btn = tk.Button(root, text="Sélectionner le fichier CSV", command=analyser_contacts)
btn.pack()

root.mainloop()
