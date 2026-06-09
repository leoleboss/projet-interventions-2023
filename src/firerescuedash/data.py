import pandas as pd


def charger_donnees(chemin_fichier):
    """
    Charge le fichier CSV dans un DataFrame pandas.
    """

    df = pd.read_csv(
        chemin_fichier,
        sep=";",
        encoding="utf-8"
    )

    return df
    