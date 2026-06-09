def top_10_departements(df):
    """
    Retourne les 10 départements avec le plus d'interventions.
    """
    resultat = (
        df[["d�partement", "total_interventions"]]
        .sort_values(by="total_interventions", ascending=False)
        .head(10)
    )
    return resultat


def grandes_categories(df):
    """
    Calcule les totaux des grandes catégories d'interventions.
    """
    categories = {
        "Incendies": df["incendies"].sum(),
        "Secours à personne": df["secours_�_personne"].sum(),
        "Accidents": df["accidents_de_circulation"].sum(),
        "Opérations diverses": df["op�rations_diverses"].sum()
    }

    return categories

    