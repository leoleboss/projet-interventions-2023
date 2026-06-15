import pandas as pd


URL_DONNEES = (
    "https://minio.lab.sspcloud.fr/"
    "leoleboss/Projet_outildedonnees/interventions2023.csv"
)


def charger_donnees():
    """
    Charge les données depuis MinIO.
    """

    df = pd.read_csv(
        URL_DONNEES,
        sep=";",
        encoding="latin1"
    )

    return df
    