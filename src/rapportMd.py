import os
import pandas as pd

# Dossiers du projet
BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # src/
DATA_DIR = os.path.join(BASE_DIR, "..", "data")         # ../data

# Fichier Excel produit par tes macros (avec les onglets d'attaque)
EXCEL_FILE = os.path.join(DATA_DIR, "logs_extraits.xlsm")

# Noms des feuilles créées par le VBA
SHEET_PS = "Onglet_PortScan"
SHEET_DOS = "Onglet_DoS"
SHEET_SSH = "Onglet_SSH_Sensible"

# Fichier de sortie Markdown
MD_FILE = os.path.join(DATA_DIR, "rapport_attaques.md")


def charger_feuilles():
    """Charge les trois feuilles d'attaque en DataFrame."""
    df_ps = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PS)
    df_dos = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_DOS)
    df_ssh = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_SSH)

    colonnes = [
        "TYPE_ATTAQUE",
        "TIME",
        "SOURCE",
        "DESTINATION",
        "PORT",
        "FLAGS",
        "LENGTH",
        "COMMENTAIRE",
    ]

    df_ps = df_ps[colonnes]
    df_dos = df_dos[colonnes]
    df_ssh = df_ssh[colonnes]

    return df_ps, df_dos, df_ssh


def generer_markdown(df_ps, df_dos, df_ssh) -> str:
    """Construit le texte Markdown du rapport."""
    total = len(df_ps) + len(df_dos) + len(df_ssh)
    nb_ps = len(df_ps)
    nb_dos = len(df_dos)
    nb_ssh = len(df_ssh)

    lignes = []

    # Titre et contexte
    lignes.append("# Rapport d’analyse de trafic réseau\n\n")
    lignes.append(
        "Ce rapport présente les activités suspectes détectées à partir du fichier "
        "tcpdump, après extraction en CSV, traitement sous Excel/VBA et génération "
        "de graphiques de synthèse.\n\n"
    )

    # Résumé global
    lignes.append("## Résumé global\n\n")
    lignes.append(f"- Nombre total de paquets marqués comme suspects : **{total}**.\n")
    lignes.append(f"- Paquets liés à un port scan : **{nb_ps}**.\n")
    lignes.append(f"- Paquets liés à un début de déni de service (DoS) : **{nb_dos}**.\n")
    lignes.append(
        f"- Paquets liés à une activité SSH sensible : **{nb_ssh}**.\n\n"
    )

    # Diagramme global
    lignes.append("### Diagramme des attaques par type\n\n")
    lignes.append("![Nombre d'attaques par type](graphb.png)\n\n")

    # =====================
    # Port scan
    # =====================
    lignes.append("## Port scan\n\n")
    lignes.append(
        "Un port scan correspond à une série de paquets TCP avec le drapeau SYN "
        "envoyés vers des ports consécutifs sur le même serveur, afin de découvrir "
        "quels ports sont ouverts.\n\n"
    )

    lignes.append("### Visualisation du port scan\n\n")
    lignes.append(
        "![Port scan : ports testés dans le temps](graphc.png)\n\n"
    )

    if nb_ps > 0:
        lignes.append("### Exemples de paquets détectés\n\n")
        lignes.append(
            "| TIME | SOURCE | DESTINATION | PORT | FLAGS | COMMENTAIRE |\n"
        )
        lignes.append(
            "|------|--------|-------------|------|-------|-------------|\n"
        )
        for _, r in df_ps.head(10).iterrows():
            lignes.append(
                f"| {r['TIME']} | {r['SOURCE']} | {r['DESTINATION']} | "
                f"{r['PORT']} | {r['FLAGS']} | {r['COMMENTAIRE']} |\n"
            )
        lignes.append("\n")
    else:
        lignes.append("Aucun comportement de port scan n’a été détecté.\n\n")

    # =====================
    # DoS
    # =====================
    lignes.append("## Déni de service (DoS)\n\n")
    lignes.append(
        "Un début de déni de service est caractérisé par un volume anormalement "
        "élevé de paquets envoyés dans un intervalle de temps très court vers la "
        "même destination.\n\n"
    )

    lignes.append("### Graphique du volume de paquets DoS\n\n")
    lignes.append(
        "![DoS : volume de paquets dans le temps](graph_dos_volume.png)\n\n"
    )

    if nb_dos > 0:
        lignes.append("### Exemples de paquets détectés\n\n")
        lignes.append(
            "| TIME | SOURCE | DESTINATION | LENGTH | COMMENTAIRE |\n"
        )
        lignes.append(
            "|------|--------|-------------|--------|-------------|\n"
        )
        for _, r in df_dos.head(10).iterrows():
            lignes.append(
                f"| {r['TIME']} | {r['SOURCE']} | {r['DESTINATION']} | "
                f"{r['LENGTH']} | {r['COMMENTAIRE']} |\n"
            )
        lignes.append("\n")
    else:
        lignes.append("Aucun comportement de déni de service n’a été détecté.\n\n")

    # =====================
    # SSH sensible
    # =====================
    lignes.append("## Activité SSH sensible\n\n")
    lignes.append(
        "L’activité SSH est considérée comme sensible lorsqu’un serveur "
        "d’administration est contacté sur un port non standard, ce qui peut "
        "signaler une tentative d’accès inhabituelle.\n\n"
    )

    if nb_ssh > 0:
        lignes.append("### Exemples de paquets détectés\n\n")
        lignes.append(
            "| TIME | SOURCE | DESTINATION | PORT | COMMENTAIRE |\n"
        )
        lignes.append(
            "|------|--------|-------------|------|-------------|\n"
        )
        for _, r in df_ssh.head(10).iterrows():
            lignes.append(
                f"| {r['TIME']} | {r['SOURCE']} | {r['DESTINATION']} | "
                f"{r['PORT']} | {r['COMMENTAIRE']} |\n"
            )
        lignes.append("\n")
    else:
        lignes.append("Aucune activité SSH sensible n’a été détectée.\n\n")

    # Conclusion
    lignes.append("## Conclusion\n\n")
    lignes.append(
        "Les comportements observés (scan de ports, pics de trafic et connexions SSH "
        "sensibles) expliquent la saturation du réseau constatée sur le site de "
        "production et justifient la mise en place de mesures de sécurité "
        "complémentaires.\n"
    )

    return "".join(lignes)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    df_ps, df_dos, df_ssh = charger_feuilles()
    contenu = generer_markdown(df_ps, df_dos, df_ssh)
    with open(MD_FILE, "w", encoding="utf-8") as f:
        f.write(contenu)
    print(f"Rapport Markdown généré : {MD_FILE}")


if __name__ == "__main__":
    main()
