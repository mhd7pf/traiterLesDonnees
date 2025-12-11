import tkinter as tk
from tkinter import filedialog
import re
from datetime import datetime, timezone, timedelta  
from zoneinfo import ZoneInfo


def traitement_dates(dates):
    """Convertit un timestamp ICS (UTC) en datetime locale Europe/Paris."""
    m = re.match(r"(\d{8}T\d{6})Z$", dates)
    if not m:
        raise ValueError(f"Format de date ICS invalide : {dates}")

    dt_utc = datetime.strptime(m.group(1), "%Y%m%dT%H%M%S").replace(tzinfo=timezone.utc)

    try:
        return dt_utc.astimezone(ZoneInfo("Europe/Paris"))
    
    except Exception:
        return dt_utc + timedelta(hours=1)

def suprimer(text):
    """Supprime les retours pliés ICS (lignes repliées)."""
    return re.sub(r"\r?\n[ \t]+", "", text)

def extract_evennements(contenu):
    """Extrait tous les événements VEVENT d'un fichier ICS."""
    evenement = re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", contenu, flags=re.S)
    parsed = []

    for block in evenement:
        block = suprimer(block)
        props = {}

        for line in block.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            props[key.strip().upper()] = value.strip()

        parsed.append(props)

    return parsed


def convert_event_to_row(props):
    """Transforme un événement ICS en liste correspondant à une ligne CSV."""

    uid = props.get("UID", "vide")
    summary = props.get("SUMMARY", "vide")
    location = props.get("LOCATION", "vide")


    if "DTSTART" in props:
        dt_start = traitement_dates(props["DTSTART"])
        date = dt_start.strftime("%d-%m-%Y")
        heure = dt_start.strftime("%H:%M")
    else:
        date = heure = "vide"

    if "DTEND" in props and "DTSTART" in props:
        dt_end = traitement_dates(props["DTEND"])
        delta = dt_end - dt_start
        minutes = int(delta.total_seconds() // 60)
        duree = f"{minutes//60:02d}:{minutes%60:02d}"
    else:
        duree = "vide"


    modalite = "vide"
    for mod in ["CM", "TD", "TP", "Proj", "DS"]:
        if re.search(rf"\b{mod}\b", summary, re.I):
            modalite = mod
            break
        if "DESCRIPTION" in props and re.search(rf"\b{mod}\b", props["DESCRIPTION"], re.I):
            modalite = mod
            break


    desc = props.get("DESCRIPTION", "")
    profs = "vide"
    groupes = "vide"

    if desc:
        items = [x.strip() for x in desc.split("\\n") if x.strip()]
        if len(items) >= 1:
            groupes = items[0]
        if len(items) >= 2:
            profs = items[1]


    location = location.replace(",", "|") if location else "vide"

    return [
        uid,
        date,
        heure,
        duree,
        modalite,
        summary,
        location,
        profs,
        groupes
    ]


def traiter_fichier_ics(chemin):
    """Lit un fichier ICS, convertit les événements et génère un CSV."""
    

    with open(chemin, encoding="utf-8") as f:
        contenu = f.read()

    evenement = extract_evennements(contenu)


    entetes = [
        "uid",
        "date",
        "heure",
        "duree",
        "modalite",
        "intitule",
        "salles",
        "profs",
        "groupes"
    ]

    lignes = []
    for e in evenement:
        lignes.append(convert_event_to_row(e))


    with open("evenements_pseudo.csv", "w", encoding="utf-8") as f:
        f.write(";".join(entetes) + "\n")
        for ligne in lignes:
            f.write(";".join(ligne) + "\n")

    return len(evenement)


def choisir_fichier():
    """Ouvre une fenêtre pour sélectionner un fichier ICS puis lance le traitement."""
    chemin_fichier = filedialog.askopenfilename(
        title="Sélectionner un fichier .ics",
        filetypes=[("Fichier ICS", "*.ics")]
    )

    if chemin_fichier:
        try:
            nb = traiter_fichier_ics(chemin_fichier)
            label_info.config(
                text=f"{nb} événements traités.\nFichier généré : evenements_pseudo.csv"
            )
        except Exception as e:
            label_info.config(text=f"Erreur : {e}")
    else:
        label_info.config(text="Aucun fichier sélectionné.")


def quitter():
    fenetre.destroy()



fenetre = tk.Tk()
fenetre.title("Convertisseur ICS → CSV")
fenetre.geometry("480x260")

btn_choisir = tk.Button(fenetre, text="Choisir un fichier .ics", command=choisir_fichier)
btn_choisir.pack(pady=20)

label_info = tk.Label(fenetre, text="Aucun fichier sélectionné.")
label_info.pack(pady=20)

btn_quitter = tk.Button(fenetre, text="Quitter", command=quitter)
btn_quitter.pack(pady=20)

fenetre.mainloop()
