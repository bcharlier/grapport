import sys
sys.path.append("..")

import grapport

# Définit le répertoire des données
grapport.set_data_dir("./")

# Charge les données des candidatures
candidature = grapport.Candidatures("encrypted_data.csv")

# Définit le rapporteur
votre_nom = "Camille Dupont"
votre_ville = "Vierzon"  # None si pas de ville
votre_signature = "signature.png"  # None si pas de signature
vos_rapports = "all"  # Liste des candidats à rapporter, peut être une liste de noms, d'index ou de numéros de candidat, ou "all" pour tous les candidats

rapporteur = grapport.Rapporteur(votre_nom,
                                 ville=votre_ville,
                                 signature=votre_signature,
                                 rapports=vos_rapports,
                                 )

# Génère les rapports
grapport.Rapport(candidature, rapporteur).generate()

