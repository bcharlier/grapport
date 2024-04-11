import os
import subprocess

import pandas as pd

import grapport.config


class Jury:
    def __init__(self, extraction_galaxie, liste_des_rapports):
        self.candidatures = self.get_candidatures(extraction_galaxie)
        self.liste_des_rapports = self.get_rapporteurs(liste_des_rapports)
        self.rapporteurs = set(self.liste_des_rapports["Nom du Rapporteur1"].unique()).union(set(self.liste_des_rapports["Nom du Rapporteur2"].unique()))

    def get_rapporteurs(self, filename):
        if filename.endswith('.csv'):
            return pd.read_csv(filename, sep=";", header=0)
        elif filename.endswith('.xls'):
            return pd.read_html(filename, header=0)[0]
        else:
            raise ValueError("Unknown file format")

    def get_candidatures(self, filename):
        if filename.endswith('.csv'):
            return pd.read_csv(filename, sep=";", header=0)
        elif filename.endswith('.xls'):
            return pd.read_excel(filename, dtype=object)
        else:
            raise ValueError("Unknown file format")

    def generate(self):
        for r in self.rapporteurs:
            df_r = self.liste_des_rapports.loc[self.liste_des_rapports['Nom du Rapporteur1'].str.contains(r) | self.liste_des_rapports['Nom du Rapporteur2'].str.contains(r), ['N° de Candidat']]
            print(r + " : " + str(len(df_r)) + " candidats")

            # Crée un répertoire pour chaque rapporteur
            current_dir = os.path.join(grapport.config.data_dir, r.replace(" ", "_"))
            os.makedirs(current_dir, exist_ok=True)
            print(current_dir)

            # Crée un fichier candidats.csv pour chaque rapporteur
            cond = self.candidatures['N° candidat'].isin(df_r['N° de Candidat'])
            self.candidatures.loc[cond, :].to_csv(os.path.join(current_dir, "candidats.csv"), sep=";", index=False)

            # Crée un fichier generate.py pour chaque rapporteur
            with open(current_dir + "/generate.py", "w") as f:
                f.write(f"""
import os, sys

sys.path.append("../../..")

import grapport

# Définit le répertoire des données
grapport.set_data_dir(os.path.join(os.path.dirname(__file__)))

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
                subprocess.run(["python", "generate.py"], cwd=current_dir, check=True)
            except subprocess.CalledProcessError as e:
                print(e)
                print(f"Erreur lors de la génération des rapports pour {r}. Veuillez lancer manuellement le script dans {current_dir}.")
