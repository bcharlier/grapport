import os.path
import random

import pandas as pd

import graport.config


class Candidatures:
    reviewer = 'Benjamin Charlier'

    def __init__(self, fname):
        self.filename = os.path.join(graport.config.data_dir, fname)
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
    candidatures = Candidatures("../data/Extraction galaxie poste 169.xls")
    print(candidatures.get_candidate(0))

    import pandas as pd


    def caesar_cipher(text):
        """
        Encrypts text using Caesar Cipher with a random shift for each word, including digits without modifying their length.

        Parameters:
        text (str): Text to be encrypted.

        Returns:
        str: Encrypted text.
        """
        words = text.split()
        encrypted_text = ""
        for word in words:
            # Generate a random shift value between 1 and 25 for each word
            shift = random.randint(1, 25)
            encrypted_word = ""
            for char in word:
                if char.isalpha():
                    shifted = ord(char) + shift
                    if char.islower():
                        if shifted > ord('z'):
                            shifted -= 26
                        encrypted_word += chr(shifted)
                    elif char.isupper():
                        if shifted > ord('Z'):
                            shifted -= 26
                        encrypted_word += chr(shifted)
                elif char.isdigit():
                    encrypted_word += str(random.randint(0, 9))  # Keep the digit unchanged
                else:
                    encrypted_word += char
            encrypted_text += encrypted_word + " "
        return encrypted_text.strip()

    df = candidatures.data.head()

    # Encrypt all text in the DataFrame using Caesar Cipher with random shift for each word and digit
    df_encrypted = df.applymap(lambda x: caesar_cipher(str(x)))

    print(df_encrypted)
    df_encrypted.to_csv("../data/encrypted_data.csv", sep=";", index=False)
