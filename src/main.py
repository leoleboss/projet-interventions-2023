from chargement import charger_donnees
from nettoyage import nettoyer_noms_colonnes, convertir_colonnes_numeriques
from analyse import top_10_departements, grandes_categories
from graphiques import graphique_top_departements, graphique_categories


# Chemins des fichiers
CHEMIN_DONNEES = "data/interventions2023.csv"


def main():
    # 1. Charger les données
    df = charger_donnees(CHEMIN_DONNEES)

    # 2. Nettoyer les données
    df = nettoyer_noms_colonnes(df)
    df = convertir_colonnes_numeriques(df)

    # 3. Faire les analyses
    top_departements = top_10_departements(df)
    categories = grandes_categories(df)

    # 4. Afficher les résultats
    print("\nTop 10 départements :")
    print(top_departements)

    print("\nGrandes catégories :")
    print(categories)

    # 5. Créer les graphiques
    graphique_top_departements(top_departements)
    graphique_categories(categories)

    print("\nProjet exécuté avec succès.")


if __name__ == "__main__":
    main()
    