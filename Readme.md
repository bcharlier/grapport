# Grapport

Générateur de rapport en **docx** ou **latex** pour les Comités de Sélection.

# installation

Dans un terminal, exécutez les commandes suivantes :
```bash
git clone https://github.com/bcharlier/grapport.git
cd grapport
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

# Démo : deux cas d'utilisation

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

# Adapter à votre cas (PR welcome)

Pour adapter le rendu à votre cas, vous pouvez ajouter un sous-répertoire dans le module[`grapport.templates`](./grapport/templates/).

# TODO (PR Welcome again...)

- [ ] documentation

# Auteur

Benjamin Charlier - [benjamin.charlier@umontpellier.fr](mailto:benjamin.charlier@umontpellier.fr)

# Licence

MIT, 2024.
