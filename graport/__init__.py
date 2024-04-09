import os

from .Candidatures import Candidatures
from .Rapport import Rapport
from .Rapporteur import Rapporteur

import graport.config

def set_data_dir(path):
    # Vérifie que le chemin existe
    if not os.path.exists(path):
        raise FileNotFoundError(f"Le chemin {path} n'existe pas.")
    graport.config.data_dir = path