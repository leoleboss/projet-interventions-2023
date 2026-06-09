import pandas as pd


def nettoyer_noms_colonnes(df):
    """
    Nettoie les noms des colonnes.
    """
    df.columns = (
        df.columns
        .str.lower()
        .str.strip()
        .str.replace(" ", "_")
    )
    return df


def convertir_colonnes_numeriques(df):
    """
    Convertit les colonnes numériques lues comme du texte.
    """
    colonnes_texte = [
        "zone",
        "r�gion",
        "num�ro",
        "d�partement",
        "cat�gorie_a"
    ]

    for colonne in df.columns:
        if colonne not in colonnes_texte:
            df[colonne] = df[colonne].astype(str)
            df[colonne] = df[colonne].str.replace(" ", "")
            df[colonne] = pd.to_numeric(df[colonne], errors="coerce")

    return df
    