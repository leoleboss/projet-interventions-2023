import matplotlib.pyplot as plt


def graphique_top_departements(top_departements):
    """
    Crée un graphique du top 10 des départements.
    """
    plt.figure(figsize=(12, 6))

    plt.bar(
        top_departements["d�partement"],
        top_departements["total_interventions"]
    )

    plt.title("Top 10 des départements avec le plus d'interventions")
    plt.xlabel("Départements")
    plt.ylabel("Nombre d'interventions")
    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.savefig("outputs/top_10_departements.png")
    plt.close()


def graphique_categories(categories):
    """
    Crée un graphique des grandes catégories.
    """
    plt.figure(figsize=(10, 6))

    plt.bar(
        categories.keys(),
        categories.values()
    )

    plt.title("Répartition des grandes catégories d'interventions")
    plt.ylabel("Nombre d'interventions")
    plt.xticks(rotation=10)
    plt.tight_layout()

    plt.savefig("outputs/categories_interventions.png")
    plt.close()
    