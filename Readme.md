# Grapport

Générateur de rapport pour les Comités de Sélection.

# installation

Dans un terminal, exécutez les commandes suivantes :
```bash
git clone https://github.com/bcharlier/grapport.git
cd grapport
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

# Deux cas d'utilisation

## Rapporteur

Dans un terminal,
```bash
cd exemple/rapporteur
python main_rapporteur.py
```

## Jury du Comité de Sélection

Dans un terminal,
```bash
cd exemple/jury
python main_jury.py
```

# Adapter à votre cas

Pour adapter le rendu à votre cas, vous pouvez modifier les fichiers [`Rapport.py`](./grapport/Rapport.py).

# TODO (PR welcome)

- [ ] créer un répertoire de templates pour différents types de rapport/institution
- [ ] refactoriser le code pour permettre l'ajout de nouveaux types de rapport (latex, docx, html, etc.)


# Auteur

Benjamin Charlier - [benjamin.charlier@umontpellier.fr](mailto:benjamin.charlier@umontpellier.fr)

# Licence

MIT, 2024.