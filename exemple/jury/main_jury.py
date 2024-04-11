import sys, os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))

import grapport

# Définit le répertoire des données
grapport.set_data_dir(os.path.join(os.path.dirname(__file__)))

# Charge les données du jury
jury = grapport.Jury("extraction galaxie.xls", "Liste_des_rapporteurs_9999999X_666.csv")

# Génère les rapports
jury.generate()
