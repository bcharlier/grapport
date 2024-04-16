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
logo = os.path.join(os.path.dirname(__file__), "unistra2.png") # None si pas de logo

def create_document(candidat, rapporteur):

    # ---------------------------------------------
    # Entête
    # ---------------------------------------------

    document = r"""
% Créée par V. Dotsenko et R. Noot à partir du fichier Word fourni par la DRH de l'Unistra 
% version 3.0, 07/04/2024 
\documentclass[a4paper]{article}
\usepackage{calc,amsmath,amssymb,amsfonts}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[french]{babel}
\usepackage{fancyhdr}
\usepackage[usenames,dvipsnames]{xcolor}
\usepackage{mdframed,graphicx}
\usepackage[top=0.2in,bottom=0.3in,hmargin=0.45in,includeheadfoot,head=27pt,headsep=0.1in,foot=0.2in,footskip=0.4in]{geometry}
\usepackage{array,supertabular,hhline,enumitem,hyperref,multirow}
\newcommand\textstyleListLabeliv[1]{\textbf{\textup{#1}}}
\newcommand\textstyleListLabelxxviii[1]{\textbf{\textup{#1}}}
\newcommand\textstyleListLabelxvi[1]{\textbf{\textup{#1}}}
\newcommand\textstyleListLabelxxxiv[1]{\textbf{#1}}
\newcommand\textstyleListLabelxxxiii[1]{\textbf{#1}}
\makeatletter
\newcommand\arraybslash{\let\\\@arraycr}
\makeatother
\fancypagestyle{FicheUnistra}{\fancyhf{}
    \fancyhead[L]{\includegraphics[height=2cm]{unistra2.png}}
  \fancyfoot[L]{}
  \renewcommand\headrulewidth{0pt}
  \renewcommand\footrulewidth{0pt}
  \renewcommand\thepage{\arabic{page}}
}
\renewcommand{\familydefault}{\sfdefault}
\pagestyle{FicheUnistra}
\setlength{\headheight}{57.8024pt}
\setlength\tabcolsep{1mm}
    """
    # ---------------------------------------------
    # "rapi","rape","candnom","candpre","bdate","situationpro","lieuex","titth","dateth","lieuth","dirth","seccnu","actrech"
    # ---------------------------------------------

    # TODO: ajouter un moyen de coder exterieur/interieur
    if not isinstance(rapporteur, list):
        rapporteur = [rapporteur, rapporteur]

    document += f"\\def\\rapp{{{rapporteur[0].nom}}}\n"
    document += f"\\def\\rape{{{rapporteur[1].nom}}}\n"
    document += f"\\def\\rebox{{$\square$}}\n"
    document += f"\\def\\rapi{{{rapporteur[1].nom}}}\n"
    document += f"\\def\\ribox{{$\square$}}\n"


    document += f"\\def\\candnom{{{candidat['Nom']}}}\n"
    document += f"\\def\\candpre{{{candidat['Prénom']}}}\n"
    document += f"\\def\\bdate{{{candidat['Né(e) le']}}}\n"
    document += f"\\def\\situationpro{{{candidat['Situation professionnelle']}}}\n"
    clexer = candidat["Lieu d'exercice"]
    document += f"\\def\\lieuex{{{clexer}}}\n"
    document += f"\\def\\titth{{{candidat['Titre thèse']}}}\n"
    document += f"\\def\\dateth{{{candidat['Date soutenance']}}}\n"
    document += f"\\def\\lieuth{{{candidat['Lieu soutenance']}}}\n"
    document += f"\\def\\dirth{{{candidat['Directeur Thèse']}}}\n"
    document += f"\\def\\seccnu{{{candidat['Section']}}}\n"
    document += f"\\def\\actrech{{{candidat['Activités recherche']}}}\n"
    aqualif = candidat["Référence qualif"].split("-")[1] if candidat["Référence qualif"] != "" else ""
    document += f"\\def\\aqualif{{{aqualif}}}\n"

    document += r"""

\begin{document}


\mdfdefinestyle{exampleu}{% 
rightline=true,innerleftmargin=10,innerrightmargin=10, frametitlerule=true,frametitlerulecolor=lightgray,backgroundcolor=lightgray, frametitlerulewidth=2pt}
\begin{mdframed}[style=exampleu]
\noindent
\textcolor{black}{Composante d'affectation: UFR de mathématique et d’informatique}

\noindent
\textcolor{black}{Section CNU: 25 \hskip 1em Poste \hskip 1em
  $\boxtimes$ MC \hskip 1em $\square$ PR \hskip 4em n${}^\circ$: 4949}

\noindent
\textcolor{black}{Intitulé du poste: Mathématiques}
\end{mdframed}

\smallskip


\noindent
\textbf{Nom du rapporteur}: \rapp % votre NOM et Prénom

\noindent
Rapporteur interne \ribox \hskip 2em Rapporteur externe \rebox 


\smallskip


\mdfdefinestyle{examplea}{% 
rightline=true,innerleftmargin=10,innerrightmargin=10, frametitlerule=true,frametitlerulecolor=NavyBlue,backgroundcolor=NavyBlue, frametitlerulewidth=2pt,fontcolor=white}
\begin{mdframed}[style=examplea]
\textbf{Prénom \& Nom du candidat ou de la candidate}: \candpre\ \candnom
% NOM Prénom 

\noindent
\textbf{Situation actuelle} ({\footnotesize en précisant le lieu d'exercice
actuel et le cas échéant, la date de nomination}): \situationpro\ à \lieuex.
% vérifier ICI la situaltion actuelle
% par exemple : postdoc à l'Université de Liège 

\end{mdframed}

\mdfdefinestyle{exampledefault}{% 
rightline=true,innerleftmargin=10,innerrightmargin=10, frametitlerule=true,frametitlerulecolor=NavyBlue, frametitlebackgroundcolor=NavyBlue, frametitlerulewidth=2pt}
\begin{mdframed}[style=exampledefault,frametitle={\textcolor{white}{\hskip -1.5em CURSUS UNIVERSITAIRE}}]
\begin{itemize}[series=listWWNumii,label=\textbullet,leftmargin=*]
\item \textbf{Cursus universitaire} ({\footnotesize discipline et  université}):
% remplir ICI le cursus universitaire
% par exemple : élève de l’ENS Ulm (2010-2014), M2 à Paris 6
% (2013-2014), doctorat à Paris 13 (2014-2017)  
\item \textbf{Thèse} ({\footnotesize titre, directeur de thèse, année
  et lieu de soutenance}): \titth, soutenue le \dateth\ à
  \lieuth\ sous la direction de \dirth. 
\item \textbf{HDR} ({\footnotesize titre, garant, année et lieu de soutenance}):
% remplir ICI si HDR
% normalement sera vide pour la plupart de personnes
\item \textbf{Section CNU}: \seccnu  
\hskip 5cm
\textbf{Date de la qualification}: \aqualif 
\end{itemize}
\end{mdframed}

\begin{mdframed}[style=exampledefault,frametitle={\textcolor{white}{\hskip -1.5em ACTIVITÉS DE RECHERCHE}}]
\begin{itemize}[resume*=listWWNumii,leftmargin=*]
\item \textbf{Domaine de spécialité et thématiques de recherche }: \actrech
% Vérifier ICI les thématiques de recherche
% une liste de mots-cles
\end{itemize}


\begin{itemize}[resume*=listWWNumii,leftmargin=*]
\item \textbf{Expérience / recherche post-doctorale} ({\footnotesize si oui, préciser le lieu et la durée}):
% remplir ICI la carrière post-doctorale
% par exemple : Cambridge (2017-2020), Université Lyon 1 (2020-2023),
% Université de Liège depuis 2023 
\end{itemize}


\begin{itemize}[resume*=listWWNumii,leftmargin=*]
\item \textbf{Publications} ({\footnotesize préciser le nombre et le rang dans la liste des auteurs et le cas échéant, préciser les
revues et/ou les maisons d'édition, la copublication avec des doctorants\ldots})
\end{itemize}
\begin{enumerate}[series=listWWNumxii,label=\arabic*.,ref=\arabic*]
\item Articles dans des revues à comité de lecture:
% remplir ICI et dans les 6 points suivants les publications etc
% par exemple : 8 articles (Math Ann, Compositio, Math Z, Ann ENS,
% IMRN, Ann ENS, J Noncomm Geom, J Lie Theory) et 2 prépublications 
\item Ouvrages: 
% nombre
\item Chapitres d'ouvrage:
% nombre
\item Direction d'ouvrages collectifs:
% nombre
\item Communications à des congrès (nationaux / internationaux):
% nombre et des précisions si des congrès prestigieux 
\item Séminaires invités dans des laboratoires:
% nombre
\item Brevets:
% s'il y a lieu
\end{enumerate}
\vfill
\begin{itemize}[resume*=listWWNumii]
\item \textbf{Contrats de recherches financés}
({\footnotesize préciser le cas échéant si responsable / partenaire, année 
d'obtention, organismes financeurs,\ldots}): 
% remplir ICI les contrats de recherche
\end{itemize}
\vfill
\begin{itemize}[resume*=listWWNumii]
\item \textbf{Participation à des activités de valorisation des
  connaissances scientifiques, diffusion de la culture  scientifique}
  ({\footnotesize e.g. communications grand public, articles de vulgarisation, etc.}):
  % remplir ICI valorisation etc.
  \vfill
\end{itemize}
\begin{itemize}[resume*=listWWNumii]
\item \textbf{Expérience de l'encadrement}
  ({\footnotesize\textit{e.g. }étudiants de master, thèse, chercheur   post-doctoral}):
% remplir ICI encadrement master, thèse etc
\end{itemize}

\end{mdframed}


\begin{mdframed}[style=exampledefault,frametitle={\textcolor{white}{\hskip -1.5emEXPERIENCES \ PEDAGOGIQUES}}]
\begin{itemize}[series=listWWNumxi,label=\textbullet,leftmargin=*]
\item \textbf{Statut} {\footnotesize(MC, ATER, chargé d'enseignement vacataire,\ldots) Le cas échéant préciser, si mobilité
(\textit{i.e.} au sein d'une autre université que celle où le doctorat a été soutenu)}:
% remplir ICI le statut actuel et mobilité après thèse
\item \textbf{Discipline d'enseignement} {\footnotesize(préciser le
  niveau: L ou M, la nature et volume d'enseignement) }: 
% remplir ICI la discipline et le niveau enseignés
\end{itemize}
\end{mdframed}

\clearpage


\begin{mdframed}[style=exampledefault,frametitle={\textcolor{white}{\hskip -1.5em
        TÂCHES D'INTÉRÊT COLLECTIF: IMPLICATION LOCALE, NATIONALE ou
        INTERNATIONALE
        \newline {\footnotesize\textcolor{white}{RQ: critères d'évaluation \ à pondérer pour les candidatures à un poste de MC
et/ou en fonction du profil de poste}}}}]
\begin{itemize}[series=listWWNumvii,label=\textbullet,leftmargin=*]
\item Responsabilités pédagogiques ({\footnotesize\textit{e.g. } responsabilité de diplôme, d'année,\ldots}):
% remplir ICI responsabilités pédagogiques
\item Responsabilités scientifiques ({\footnotesize\textit{e.g. }responsabilité de contrat, responsabilité
d'équipe\ldots}): 
% ICI responsabilités scientifiques
\item Responsabilités administratives ({\footnotesize \textit{e.g. } mandat au sein de la gouvernance de
l'université, au sein d'une composante, membre d'un conseil\ldots}): 
% ICI responsabilités administratives
\item Relations avec le monde socio-économique ({\footnotesize\textit{e.g. }partenariats avec des entreprises, des institutions
publiques, le monde associatif,\ldots}): 
% ICI relations avec le monde socio-économique
\item Autres: 
% ICI autre tâches d'intérêt collectif
\end{itemize}
\end{mdframed}

\noindent
{\footnotesize\textcolor{black}{En référence à la labellisation HRS4R obtenue par l'Unistra en 2017, il est
demandé à chaque rapporteur de porter un avis sur la }\textbf{qualité} du dossier
de candidature et sur l'\textbf{adéquation} au profil de
poste.}

\begin{flushleft}
\tablefirsthead{}
\tablehead{}
\tabletail{}
\tablelasttail{}
\begin{supertabular}{|m{0.9in}|m{2.1in}|m{0.51155984in}|m{0.6101598in}|m{0.6115598in}|m{2.0858598in}|}
\hhline{~~----}
\multicolumn{2}{m{3.1in}|}{\footnotesize{\textbf{\textcolor{gray}{Réservé}}\textcolor{gray}{: candidature en-deçà de ce qui
      peut être attendu}}
  \textbf{\textcolor{gray}{Favorable}}\textcolor{gray}{: candidature correspondant à ce
qui est attendu}

\textbf{\textcolor{gray}{Très favorable}}\textcolor{gray}{: candidature au-delà de ce qui peut être attendu}} &
\centering \textcolor[HTML]{0070C0}{Avis réservé} &
\centering \textcolor[HTML]{0070C0}{Avis favorable} &
\centering \textcolor[HTML]{0070C0}{Avis très favorable} &
\centering\arraybslash \textcolor[HTML]{0070C0}{Commentaires éventuels}\\\hline
\multirow{3}{0.9in}{\centering \textbf{\textcolor[HTML]{0070C0}{Avis
      sur la qualité}}} &
\textcolor[HTML]{0070C0}{de la qualification universitaire} & ~ & ~ & ~ & ~ \\
% Cocher ICI avis qualification : & réservé & favorable  & très favorable & commentaires
\hhline{~-----}
&
\textcolor[HTML]{0070C0}{du dossier Enseignement} & ~ & ~ & ~ & ~ \\
% Cocher ICI avis enseignement : & réservé & favorable  & très  favorable & commentaires
\hhline{~-----}
 &
\textcolor[HTML]{0070C0}{du dossier Recherche}\textcolor[HTML]{0070C0}{\textsuperscript{(1)}}
& ~ & ~ & ~ & ~ \\
% Cocher ICI avis recherche : & réservé & favorable  & très  favorable & commentaires
\hhline{~-----}
 &
\textcolor[HTML]{0070C0}{de l'investissement institutionnel}\textcolor[HTML]{0070C0}{\textsuperscript{(2)}}
& ~ & ~ & ~ & ~ \\
% Cocher ICI avis investissement : & réservé & favorable  & très favorable & commentaires
\hline
\multirow{2}{0.9in}{\centering \textbf{\textcolor[HTML]{0070C0}{Avis sur
l'adéquation}}} &
\textcolor[HTML]{0070C0}{au profil Enseignement} & ~ & ~ & ~ & ~ \\
% Cocher ICI adéquation profil recherche : & réservé & favorable  & très favorable & commentaires
\hhline{~-----}
 &
\textcolor[HTML]{0070C0}{au profil Recherche} & ~ & ~ & ~ & ~ \\
% Cocher ICI adéquation profil recherche : & réservé & favorable  & très favorable & commentaires
\hline
\end{supertabular}
\end{flushleft}
{\footnotesize
\noindent
\begin{enumerate}[series=listWWNumxxx,label=\textstyleListLabelxxxiv{(\arabic*)},ref=\arabic*]
\item \textcolor{gray}{L'appréciation du dossier ne doit pas consister exclusivement à une
évaluation bibliométrique. \ La qualité du dossier de publications doit être appréciée non seulement quantitativement
mais également qualitativement. Les co-publications peuvent être également valorisées. }
\item \textcolor{gray}{L'importance de ce critère d'évaluation comparativement
aux autres critères peut être pondérée pour les candidatures à un poste de~MC.}
\end{enumerate}}
\begin{flushleft}
\tablefirsthead{}
\tablehead{}
\tabletail{}
\tablelasttail{}
\begin{supertabular}{|m{3.6622598in}|m{0.21635985in}|m{0.21635985in}|m{0.31505984in}|m{0.31505984in}|m{1.9879599in}|}
\hhline{~-----}
\multicolumn{1}{m{3.6622598in}|}{\textcolor{gray}{{\footnotesize Les rapporteurs peuvent appuyer leur appréciation globale en prenant
en compte la présence -ou non- des éléments ci-dessous (lorsque cela s'avère pertinent). }}} &
\multicolumn{2}{m{0.5114598in}|}{\centering \textcolor[HTML]{0070C0}{Oui}\textcolor[HTML]{0070C0}{\textsuperscript{(1)}}} &
\centering \textcolor[HTML]{0070C0}{Non} &
\centering \textcolor[HTML]{0070C0}{N.P} &
\centering\arraybslash \textcolor[HTML]{0070C0}{Commentaires éventuels}\\\hline
 &
\centering \textcolor[HTML]{0070C0}{+} &
\centering \textcolor[HTML]{0070C0}{++} &
 &
 &
\\\hline
\textcolor[HTML]{0070C0}{Proposition d'un programme d'enseignements } & ~ & ~ & ~ & ~ & ~ \\
% cocher ICI programme d'enseignements & oui+ & oui++ & non & NP & commentaires
\hline
\textcolor[HTML]{0070C0}{Proposition d'un projet d'intégration recherche} & ~ & ~ & ~ & ~ & ~\\
% cocher ICI programme de recherche & oui+ & oui++ & non & NP & commentaires
\hline
\textcolor[HTML]{0070C0}{Mobilité nationale } & ~ & ~ & ~ & ~ & ~ \\
% cocher ICI mobilité nationale & oui+ & oui++ & non & NP & commentaires
\hline
\textcolor[HTML]{0070C0}{Mobilité internationale} & ~ & ~ & ~ & ~ & ~ \\
% cocher ICI mobilité internationale & oui+ & oui++ & non & NP & commentaires
\hline
\textcolor[HTML]{0070C0}{Parcours multidimensionnel {\footnotesize(interruption,
chang\textcolor[HTML]{0070C0}{\textsuperscript{t}} \textcolor[HTML]{0070C0}{ professionnel,\ldots})}} & ~ & ~ & ~ & ~X & ~ \\
% cocher ICI parcours multidimensionnel !??! & oui+ & oui++ & non & NP & commentaires
% Choix NP précoché
\hline
\end{supertabular}
\end{flushleft}
{\footnotesize
\begin{enumerate}[series=listWWNumxxix,label=\textstyleListLabelxxxiii{(\arabic*)},ref=\arabic*]
\item \textcolor{gray}{+: élément présent mais peu développé - ++: élément présent, bien développé et étayé}
\item \textcolor{gray}{N.P.: non pertinent}
\end{enumerate}}


\begin{mdframed}[style=exampledefault,frametitlealignment={\hspace*{0.20\linewidth}},frametitle={\textcolor{white}{APPRECIATION GLOBALE DU DOSSIER DE CANDIDATURE\newline {\footnotesize\phantom{AAA}
Faire ressortir les points forts et les points faibles du dossier et l'adéquation
(ou non) au profil du poste (enseignement et recherche).\newline 
Le président du comité de sélection pourra reprendre ces différents éléments lors de la rédaction des
avis individuels et/ou de l'avis unique~motivé}}}]
  \vbox to 17ex{\parindent=0pt
% Remplir ICI appréciation globale c'est la partie la plus importante

\vfill }
\end{mdframed}



{\footnotesize
\begin{flushleft}
\tablefirsthead{}
\tablehead{}
\tabletail{}
\tablelasttail{}
\begin{supertabular}{|m{1.1976599in}|m{0.31435984in}|m{0.70875984in}|m{1.1018599in}|m{0.21635985in}|m{0.90595984in}|m{1.2997599in}|m{0.31435984in}|}
\hhline{--~--~--}
\centering Avis favorable pour une audition &
~
% cocher ICI si favorable audition (=é valuation A)
 &
~
 &
\centering Avis à discuter &
~
% cocher ICI si audition à discuter (= B) ; cocher aussi une raison ci-dessous
 &
~
 &
\centering Avis défavorable à l'audition &
~
% cocher ICI si défavorable audition (= C) ; cocher aussi une raison ci-dessous
\\\hhline{--~--~--}
\end{supertabular}
\end{flushleft}
% vous pouvez remplacer $\square$ par $\boxtimes$ dans les cas échéant

\noindent
\hspace*{6.2cm} $\square$ Adéquation partielle profil enseignement
\hspace*{0.7cm} $\square$ Inadéquation profil enseignement
% ICI justification évaluation B/C

\noindent
\hspace*{6.2cm}  $\square$ Adéquation partielle profil recherche
\hspace*{1.2cm} $\square$ Inadéquation profil recherche
% et/ou ICI

\noindent
\hspace*{6.2cm}  $\square$ Avis réservé sur les activités enseignement
\hspace*{0.4cm}\ $\square$ Avis défavorable sur activités enseignement
% et/ou ICI

\noindent
\hspace*{6.2cm}  $\square$ Avis réservé sur les activités recherche
\hspace*{0.9cm} $\square$ Avis défavorable sur activités recherche
% et/ou ICI


\vfill
"""
    # ---------------------------------------------
    # Signature
    # ---------------------------------------------
    document += f"""

    \\vspace*{{1cm}}

    Fait le {datetime.datetime.now().strftime("%d %B %Y")} à {rapporteur[0].ville if rapporteur[0].ville is not None else " "}  \n\\newline\n\n 
    """
    document += "\\hfill\\begin{minipage}{0.3\\textwidth}\n"
    document += (
                    f"\\includegraphics[scale=.7]{{{rapporteur[0].signature}}}" if rapporteur[0].signature is not None else "% REMPLIR ICI") + "\n\n"
    document += rapporteur[0].nom
    document += "\\end{minipage}\n"

    document += "\\end{document}"

    return document

def save(document, filename):
    with open(filename, 'w') as f:
        f.write(document)

    # save the logo to the same directory
    import shutil
    shutil.copy(logo, os.path.join(os.path.dirname(filename), os.path.basename(logo)))





if __name__ == "__main__":
    candidat = grapport.Candidatures("../exemple/jury/extraction galaxie.xls").get_candidate(1)
    rapporteur_ext = grapport.Rapporteur("Camille Dupont")
    rapporteur_int = grapport.Rapporteur("Jean Dupuis")

    print(create_document(candidat, [rapporteur_ext, rapporteur_int]))
    save(create_document(candidat, [rapporteur_ext, rapporteur_int]), "rapport-BERNARD_MAXIME-Camille_Dupont.tex")
