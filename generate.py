import graport

# Définit le répertoire des données
graport.set_data_dir("./exemple")

# Charge les données des candidatures
candidature = graport.Candidatures("encrypted_data.csv")

# Définit le rapporteur
votre_nom = "Camille Dupont"
votre_ville = "Vierzon"   # None si pas de ville
votre_signature = "signature.png"  # None si pas de signature

rapporteur = graport.Rapporteur(votre_nom,
                                ville=votre_ville,
                                signature=votre_signature,
                                rapports=[1, 2, 3]
                                )

# Génère les rapports
graport.Rapport(candidature, rapporteur).generate()

