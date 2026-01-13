import re
import csv
import os
import tkinter as tk
from tkinter import filedialog

# ==========================================================
# Expressions régulières pour le format tcpdump
# ==========================================================

REGEX_LIGNE = re.compile(
    r"(?P<time>\d{2}:\d{2}:\d{2}\.\d+)\s+"
    r"IP\s+(?P<src>.+?)\s+>\s+(?P<dst>.+?):\s+"
    r"Flags\s+\[(?P<flags>[^\]]+)\].*?"
    r"length\s+(?P<length>\d+)"
)

REGEX_PORT = re.compile(r"(.+)\.(\d+)$")

# ==========================================================
# Fonctions utilitaires
# ==========================================================

def split_ip_port(champ):
    """
    Sépare un champ de type 'hote.port'
    Retourne (hote, port) ou (champ, 'vide')
    """
    if not champ:
        return "vide", "vide"

    match = REGEX_PORT.match(champ)
    if match:
        return match.group(1), match.group(2)

    return champ, "vide"

# ==========================================================
# Extraction des logs vers CSV
# ==========================================================

def extraire_logs(fichier_txt, fichier_csv):
    with open(fichier_txt, "r", encoding="utf-8", errors="ignore") as fin, \
         open(fichier_csv, "w", newline="", encoding="utf-8") as fout:

        writer = csv.writer(fout, delimiter=";")
        writer.writerow([
            "TIME",
            "SOURCE",
            "DESTINATION",
            "PORT",
            "FLAGS",
            "LENGTH"
        ])

        for ligne in fin:
            # Ignorer les lignes hexadécimales
            if ligne.startswith("0x"):
                continue

            m = REGEX_LIGNE.search(ligne)
            if not m:
                continue

            src_brut = m.group("src")
            dst_brut = m.group("dst")

            # Extraction ports
            src_clean, port_src = split_ip_port(src_brut)
            dst_clean, port_dst = split_ip_port(dst_brut)

            # Priorité au port destination, sinon source
            port = port_dst if port_dst != "vide" else port_src

            writer.writerow([
                m.group("time"),
                src_clean,
                dst_clean,
                port,
                m.group("flags"),
                m.group("length")
            ])

# ==========================================================
# Interface graphique (IDENTIQUE AU ICS)
# ==========================================================

def choisir_fichier():
    chemin_fichier = filedialog.askopenfilename(
        title="Sélectionner un fichier de logs",
        filetypes=[("Fichier texte", "*.txt"), ("Tous les fichiers", "*.*")]
    )

    if not chemin_fichier:
        label_info.config(text="Aucun fichier sélectionné.")
        return

    # Création automatique du dossier data
    dossier_data = os.path.join(os.path.dirname(__file__), "..", "data")
    os.makedirs(dossier_data, exist_ok=True)

    chemin_csv = os.path.join(dossier_data, "logs_extraits.csv")

    try:
        extraire_logs(chemin_fichier, chemin_csv)
        label_info.config(
            text="Extraction terminée.\nFichier généré : logs_extraits.csv"
        )
    except Exception as e:
        label_info.config(text=f"Erreur : {e}")

def quitter():
    fenetre.destroy()

# ============================================================
# FENÊTRE PRINCIPALE
# ============================================================

fenetre = tk.Tk()
fenetre.title("Export Logs Réseau → CSV")
fenetre.geometry("520x260")

btn_choisir = tk.Button(
    fenetre,
    text="Choisir un fichier de logs",
    command=choisir_fichier
)
btn_choisir.pack(pady=20)

label_info = tk.Label(
    fenetre,
    text="Aucun fichier sélectionné."
)
label_info.pack(pady=20)

btn_quitter = tk.Button(
    fenetre,
    text="Quitter",
    command=quitter
)
btn_quitter.pack(pady=20)

fenetre.mainloop()
# ============================================================