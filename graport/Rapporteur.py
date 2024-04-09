import os.path

import graport.config

class Rapporteur:
    """
    Classe Rapporteur.  Un rapporteur est une personne qui rédige des rapports.

    Attributs:
    nom: str
    ville: str
    signature: str ou None (par défaut) fichier image de la signature
    rapports: list de int ou str (par défaut vide)
    """

    def __init__(self, nom, ville="", signature=None, rapports=[], type="index"):
        self.nom = nom
        self.ville = ville
        self.type = type

        # vérifie le chemin de l'image de la signature
        if signature is not None:
            if not os.path.exists(os.path.join(graport.config.data_dir, signature)):
                raise FileNotFoundError(f'Le fichier {signature} n\'existe pas.')

        self.signature = os.path.join(graport.config.data_dir, signature) if signature is not None else None
        self.rapports = rapports


    def report(self, message):
        print(f'{self.nom}: {message}')


if __name__ == '__main__':
    sig_path = os.path.join(graport.config.data_dir, 'signature.png')

    rapporteur = Rapporteur("Camille Dupont", ville="Montpellier", signature=sig_path, rapports=[1, 2, 3])


