import sys, os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))

import grapport

# Définit le répertoire des données
grapport.set_data_dir(os.path.dirname(__file__))

# Charge les données du jury : template="um_latex" ou template="um_docx"
jury = grapport.Jury("extraction galaxie.xls",
                     "Liste_des_rapporteurs_9999999X_666.xls",
                     template="unistra_latex")

# Génère les rapports
jury.generate(overwrite=False)
