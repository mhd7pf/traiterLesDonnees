import tkinter as tk
from tkinter import filedialog
import re

# ============================================================
# CONSTANTES
# ============================================================

# Champs réels du format ICS (ordre du CSV)
ENTETES_ICS = [
    "BEGIN",
    "DTSTAMP",
    "DTSTART",
    "DTEND",
    "SUMMARY",
    "LOCATION",
    "DESCRIPTION",
    "UID",
    "CREATED",
    "LAST-MODIFIED",
    "SEQUENCE",
    "END"
]


# ============================================================
# FONCTIONS DE TRAITEMENT ICS
# ============================================================

def nettoyer_lignes_repliees(contenu):
    """
    Supprime les retours à la ligne suivis d'espaces ou de tabulations.
    (spécificité du format ICS)
    """
    return re.sub(r"\r?\n[ \t]+", "", contenu)


def extraire_evenements(contenu):
    """
    Extrait tous les blocs VEVENT du fichier ICS.
    Chaque événement est stocké sous forme de dictionnaire.
    """
    blocs = re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", contenu, flags=re.S)
    evenements = []

    for bloc in blocs:
        bloc = nettoyer_lignes_repliees(bloc)
        donnees = {}

        for ligne in bloc.splitlines():
            if ":" not in ligne:
                continue

            cle, valeur = ligne.split(":", 1)
            donnees[cle.strip().upper()] = valeur.strip()

        evenements.append(donnees)

    return evenements


def evenement_vers_ligne_csv(evenement):
    """
    Convertit un événement ICS en ligne CSV
    en respectant STRICTEMENT les champs ICS.
    """
    ligne = []

    for champ in ENTETES_ICS:
        if champ == "BEGIN":
            ligne.append("VEVENT")

        elif champ == "END":
            ligne.append("VEVENT")

        else:
            valeur = evenement.get(champ, "vide")

            # Nettoyage léger de la description (sans interprétation)
            if champ == "DESCRIPTION":
                valeur = valeur.replace("\\n", " ").strip()

            ligne.append(valeur)

    return ligne


def traiter_fichier_ics(chemin_ics):
    """
    Lit un fichier ICS et génère un CSV fidèle
    aux noms et aux valeurs du format ICS.
    """
    with open(chemin_ics, encoding="utf-8") as f:
        contenu = f.read()

    evenements = extraire_evenements(contenu)

    with open("evenements_ics.csv", "w", encoding="utf-8") as f:
        # Écriture des en-têtes ICS
        f.write(";".join(ENTETES_ICS) + "\n")

        # Écriture des événements
        for ev in evenements:
            ligne = evenement_vers_ligne_csv(ev)
            f.write(";".join(ligne) + "\n")

    return len(evenements)


# ============================================================
# INTERFACE GRAPHIQUE (Tkinter)
# ============================================================

def choisir_fichier():
    """
    Ouvre une boîte de dialogue pour choisir un fichier ICS
    puis lance la conversion en CSV.
    """
    chemin = filedialog.askopenfilename(
        title="Sélectionner un fichier .ics",
        filetypes=[("Fichier ICS", "*.ics")]
    )

    if chemin:
        try:
            nb = traiter_fichier_ics(chemin)
            label_info.config(
                text=f"{nb} événements exportés\nFichier créé : evenements_ics.csv"
            )
        except Exception as e:
            label_info.config(text=f"Erreur : {e}")
    else:
        label_info.config(text="Aucun fichier sélectionné.")


def quitter():
    """Ferme l'application."""
    fenetre.destroy()


# ============================================================
# FENÊTRE PRINCIPALE
# ============================================================

fenetre = tk.Tk()
fenetre.title("Export ICS → CSV (champs réels)")
fenetre.geometry("520x260")

btn_choisir = tk.Button(
    fenetre,
    text="Choisir un fichier .ics",
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
