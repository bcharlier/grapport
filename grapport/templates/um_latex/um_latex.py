"""
This module contains the template for the document. It creates a Word document with the information of a candidate.
It should contain the following functions:

- create_document: creates the document
- save: saves the document
- extension: extension of the file
- logo: logo of the institution
"""
import datetime, os

import grapport
from grapport.utils import calculate_age

import locale
locale.setlocale(locale.LC_TIME, "")

extension = "tex"
logo = os.path.join(os.path.dirname(__file__), "logo.png") # None si pas de logo

def create_document(candidat, rapporteur):
    # ---------------------------------------------
    # Entête
    # ---------------------------------------------

    document = """
\\documentclass[leqno,oneside,12pt,a4paper]{amsart}
\\usepackage[a4paper, total={160mm,245mm}, left=25mm, top=25mm]{geometry} 
\\usepackage{times}
\\usepackage[T1]{fontenc}
\\usepackage[utf8]{inputenc}
\\usepackage{graphicx}
\\usepackage[french]{babel}
\\usepackage{amssymb,hyperref,amsmath}


\\newcommand*\\acocher{\\quitvmode{\\fboxsep0pt \\fboxrule0.8pt \\fbox{\\vrule height1.8ex depth.1ex width0pt \\vrule height0pt depth0pt width1.9ex }}}
\\newcommand*\\cochee{\\quitvmode{\\fboxsep0pt \\fboxrule0.8pt \\fbox{\\vrule height1.8ex depth.1ex width0pt \\vrule height0pt depth0pt width0pt \\begin{large}$\\times$\\end{large}}}}

\\parindent 0pt

\\newenvironment{absolutelynopagebreak}
  {\\par\\nobreak\\vfil\\penalty0\\vfilneg
   \\vtop\\bgroup}
  {\\par\\xdef\\tpd{\\the\\prevdepth}\\egroup
   \\prevdepth=\\tpd}

\\begin{document}
"""
    # ---------------------------------------------
    # Titre
    # ---------------------------------------------

    document += f"\\includegraphics[scale=1]{{{logo}}}"
    document += """
\\vspace{1cm}
\\begin{center}
{\\Large\\sffamily\\bfseries RAPPORT SUR DOSSIER DE CANDIDATURE}

\\smallskip
\\end{center}\n \\vspace{1cm}
"""

    document += """
% ---------------------------------------------
% Rapporteur
% ---------------------------------------------\n\n
"""

    document += "\n\n"
    document += f"\\section*{{N° Galaxie du Poste: {candidat['Référence GALAXIE']}}}\n\n"
    document += f"\\textbf{{Nom de la rapporteuse:}} {rapporteur.nom}\n\\newline\n\n"
    document += f"\\textbf{{Corps :}} {candidat['Corps']}\n\\hspace{{5mm}}\n"
    document += f"\\textbf{{Section :}} {candidat['Section']}/{candidat['Autre section']}\n\\hspace{{5mm}}\n" if candidat['Autre section'] != "" else f"\\textbf{{Section :}} {candidat['Section']}\n\\hspace{{5mm}}\n"
    document += f"\\textbf{{Profil :}} {candidat['Profil J.o']}\n\\newline\n\n"

    document += """
% ---------------------------------------------
% Candidat
% ---------------------------------------------\n\n
"""

    cnusage = candidat["Nom d'usage ou marital"]
    clexer = candidat["Lieu d'exercice"]

    document += f"\\section*{{Candidat N° {candidat['N° candidat']}}}\n\n"

    document += f"\\textbf{{Nom:}} {candidat['Nom']} \n\hspace{{5mm}}\n \\textbf{{Nom d'usage:}} {cnusage} \n\hspace{{5mm}}\n \\textbf{{Prénom:}} {candidat['Prénom']}\n\\newline\n\n"
    document += f"\\textbf{{Né(e) le:}} {candidat['Né(e) le']} ({ calculate_age(candidat['Né(e) le'])} ans) \\hspace{{1cm}}"
    document += f"\\textbf{{à:}} {candidat['Lieu de naissance']}\n\\newline\n\n"

    document += f"\\textbf{{Situation professionnelle:}} {candidat['Situation professionnelle']}"
    document += f"\n\\hspace{{1cm}}\n \\textbf{{Lieu d'exercice:}} {clexer}\\newline\n\n"

    document += f"\\textbf{{Qualification:}} {candidat['N° de qualif']}\n\\newline\n\n"

    document += """
% ---------------------------------------------
% Diplôme et formation
% ---------------------------------------------\n\n
"""

    document += f"\\section*{{Formation et expérience professionnelle}}\n\n"

    document += f"\\textbf{{Titre:}} {candidat['Titres']}\n\\newline\n\n"
    document += f"\\textbf{{Sujet de Thèse:}} {candidat['Titre thèse']}\n\\newline\n\n"

    document += f"\\textbf{{Date de soutenance:}} {candidat['Date soutenance']} \n\hspace{{1cm}}\n"
    document += f"\\textbf{{Etablissement:}} {candidat['Lieu soutenance']}\n\\newline\n\n"

    document += f"\\textbf{{Directeur de recherche:}} {candidat['Directeur Thèse']}\n\\newline\n\n"
    document += f"\\textbf{{Jury:}} {candidat['Jury']}\\newline\n\n"

    document += "\\textbf{Autres diplômes:}" + (f"{candidat['Autres diplômes']}" if candidat['Autres diplômes'] != "" else "" )+ "\n\\newline\n\n"


    document += f"\\textbf{{Expérience professionnelle en recherche:}}  % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Expérience professionnelle en enseignement:}}  % REMPLIR ICI \n\\newline\n\n"

    document += """
% ---------------------------------------------
% Expérience professionnelle
% ---------------------------------------------\n\n
"""

    document += "\\section*{Activités d'enseignement}\n\n"

    document += f"\\textbf{{Enseignements dispensés:}} {candidat['Activités enseignement']} \n\\newline\n\n"
    document += f"\\textbf{{Responsabilités exercées dans la formation:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Activités de vulgarisation ou de diffusion:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Remarques éventuelles:}} % REMPLIR ICI SI NECESSAIRE \n\\newline\n\n"

    document += "\\section*{Activités de recherche}\n\n"
    document += f"\\textbf{{Thématique de recherche:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Publications originales:}} {candidat['Activités recherche']} \n\\newline\n\n"
    document += f"\\textbf{{Autres publications:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Rayonnement:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Description des résultats obtenus:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Projets de recherche:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Insertion dans le laboratoire:}} % REMPLIR ICI \n\\newline\n\n"

    document += "\\section*{Encadrement}\n\n"
    document += f"\\textbf{{Encadrement de thèses:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Autres activités d'encadrement:}} % REMPLIR ICI \n\\newline\n\n"

    document += "\\section*{Responsabilités collectives:}\n\n"
    document += f"\\textbf{{Responsabilités \\og{{}} administratives \\fg{{}}:}} {candidat['Activités administratives']}\n\\newline\n\n"
    document += f"\\textbf{{Participation à des instances d'évaluation:}} % REMPLIR ICI \n\\newline\n\n"
    document += f"\\textbf{{Autres:}} % REMPLIR ICI \n\\newline\n\n"

    document += """
% ---------------------------------------------
% Proposition et avis
% ---------------------------------------------\n\n
"""
    document += f"\\section*{{Proposition et avis}}\n\n"

    document += f"\\textbf{{Adéquation au profil:}} % REMPLIR ICI \n\\smallskip\n\n"
    document += f"\\textbf{{Points forts:}} % REMPLIR ICI \n\\smallskip\n\n"
    document += f"\\textbf{{Points faibles:}} % REMPLIR ICI \n\\smallskip\n\n"

    document += """
    
\\vspace*{1cm}

\\begin{absolutelynopagebreak}
\\hrule

\\bigskip

\\begin{center}
% Changer \\acocher par \\cochee pour avoir une jolie croix dans la case
\\hspace{2cm} \\textbf{OUI} \\ \\acocher 
\\hspace{2cm} \\textbf{NON} \\ \\acocher
\\hspace{2cm} \\textbf{A DISCUTER} \\ \\acocher
\end{center}

\\bigskip

\\hrule 
\\end{absolutelynopagebreak}   
    """

    document += f"""
% ---------------------------------------------
% Signature
% ---------------------------------------------\n\n
    """

    document += f"""

\\vspace*{{1cm}}

Fait le {datetime.datetime.now().strftime("%d %B %Y")} à {rapporteur.ville if rapporteur.ville is not None else " "}  \n\\newline\n\n 
"""
    document += "\\hfill\\begin{minipage}{0.3\\textwidth}\n"
    document += (f"\\includegraphics[scale=.7]{{{rapporteur.signature}}}" if rapporteur.signature is not None else "% REMPLIR ICI") + "\n\n"
    document += rapporteur.nom
    document += "\\end{minipage}\n"

    document += "\\end{document}"

    return document

def save(document, filename):

    # write the document to a file
    with open(filename, 'w') as f:
        f.write(document)

    # save the logo to the same directory
    import shutil
    shutil.copy(logo, os.path.join(os.path.dirname(filename), os.path.basename(logo)))


if __name__ == "__main__":
    candidat = grapport.Candidatures("../exemple/jury/extraction galaxie.xls").get_candidate(0)
    rapporteur = grapport.Rapporteur("Camille Dupont")

    print(create_document(candidat, rapporteur))
    save(create_document(candidat, rapporteur), "test.tex")

