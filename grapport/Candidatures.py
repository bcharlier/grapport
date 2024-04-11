import os.path
import random

import pandas as pd

import grapport.config


class Candidatures:

    def __init__(self, fname):
        self.filename = os.path.join(grapport.config.data_dir, fname)
        self.data = self.clean_data(self.get_data())

    def get_data(self):
        if self.filename.endswith('.csv'):
            return pd.read_csv(self.filename, sep=";", header=0)
        elif self.filename.endswith('.xls'):
            return pd.read_excel(self.filename, dtype=object)
        else:
            raise ValueError("Unknown file format")

    def clean_data(self, df):
        df = df.fillna('')
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        return df

    def get_candidate(self, id, type="index"):
        if isinstance(id, int) and type == "index":
            return self.data.iloc[id, :]
        elif isinstance(id, int) and type == "number":
            return self.data[self.data["N° candidat"] == id, :]
        elif isinstance(id, list):
            return self.data[self.data["Nom"] == id, :]
        else:
            raise ValueError("Unknown type")





if __name__ == '__main__':
    grapport.set_data_dir(os.path.join(os.path.dirname(__file__), '..', "exemple", "rapporteur"))
    candidatures = Candidatures("extraction galaxie.xls")
    print(candidatures.get_candidate(0))
