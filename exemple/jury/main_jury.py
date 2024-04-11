import os.path
import subprocess
import pandas as pd
import grapport

# Rapporteurs
df = pd.read_csv(os.path.join(os.path.dirname(__file__), "Liste_des_rapporteurs_9999999X_666.csv"), header=0, sep=";")

rapporteurs = set(df["Nom du Rapporteur1"].unique()).union(set(df["Nom du Rapporteur2"].unique()))

# Candidatures
df_c = pd.read_excel(os.path.join(os.path.dirname(__file__), "extraction galaxie.xls"), header=0)

for r in rapporteurs:

    # get the list of candidates for each rapporteur:
    df_r = df.loc[df['Nom du Rapporteur1'].str.contains(r) | df['Nom du Rapporteur2'].str.contains(r), ['N° de Candidat']]
    print(r + " : " + str(len(df_r)) + " candidats")

    # create a directory for each rapporteur
    current_dir = os.path.join(os.path.dirname(__file__), r.replace(" ", "_"))
    os.makedirs(current_dir, exist_ok=True)
    cond = df_c['N° candidat'].isin(df_r['N° de Candidat'])
    df_c.loc[cond, :].to_csv(os.path.join(current_dir, "candidats.csv"), sep=";", index=False)

    # génération des scripts de fiches de rapport
    with open(os.path.join(current_dir, "generate.py"), "w") as f:
        f.write(f"""
import sys

sys.path.append("..")

import grapport

# Définit le répertoire des données
grapport.set_data_dir("./")

# Charge les données des candidatures
candidature = grapport.Candidatures("candidats.csv")

# Définit le rapporteur
votre_nom = "{r.title()}"
votre_ville = None  # None si pas de ville ou une str
votre_signature = None  # None si pas de signature ou "signature.png"
vos_rapports = "all"

rapporteur = grapport.Rapporteur(votre_nom,
                                 ville=votre_ville,
                                 signature=votre_signature,
                                 rapports=vos_rapports,
                                 )

# Génère les rapports
grapport.Rapport(candidature, rapporteur).generate()
    """)
    
    # lance la génération des fiches de rapport
    try:
        subprocess.run(["python", "generate.py"], cwd=current_dir)
    except:
        print(f"Error while generating the reports of {r.title()}. Run the script {os.path.join(current_dir, 'generate.py')} manually.")
        break