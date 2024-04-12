import datetime
import os.path



import grapport
from grapport.utils import calculate_age, short_path


class Abstract_Rapport:
    """
    Classe Rapport.  Un rapport est un document qui contient des informations sur un candidat.
    Lit les données d'un candidat et les intègre dans un document Word.

    Attributs:
    candidatures: Candidatures instance
    rapporteur: Rapporteur instance

    Méthodes:
    generate: Génère un rapport pour chaque poste.
    """

    extension = None # extension of the file


    def __init__(self, candidatures, rapporteur, template="um_docx"):
        self.candidatures = candidatures
        self.rapporteur = rapporteur
        if self.rapporteur.rapports == "all":
            self.rapporteur.rapports = list(candidatures.data.index)

        self.filename_creator = lambda candidat: os.path.join(grapport.config.data_dir,
                                                                f"rapport-{candidat['Nom'].replace(' ', '_')}_{candidat['Prénom'].replace(' ', '_')}-{self.rapporteur.nom.replace(' ', '_')}.{self.extension}}}")


    def generate(self):

        for id in self.rapporteur.rapports:
            candidat = self.candidatures.get_candidate(id, type=self.rapporteur.type)
            print(f"Generating report for candidate {candidat['Nom']} {candidat['Prénom']}...", end="")

            doc = self.create_document(candidat, self.rapporteur)
            print('done. Saving... ', end='')

            filename = self.filename_creator(candidat)
            self.save(doc, filename)
            print(f"Saved to {short_path(filename)}.")

    @staticmethod
    def create_document(candidat, rapporteur):
        pass
    @staticmethod
    def save(document, filename):
        pass



def Rapport(candidatures, rapporteur, template="um_docx"):
    Rapp =  Abstract_Rapport(candidatures, rapporteur)

    # check if there is a submodude called template in grapport.templates
    if hasattr(grapport.templates, template):
        Rapp.create_document = getattr(grapport.templates, template).create_document
        Rapp.extension = getattr(grapport.templates, template).extension
        Rapp.save = getattr(grapport.templates, template).save

    else:
        raise ValueError(f"Template {template} not found.")

    return Rapp