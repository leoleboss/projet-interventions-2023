from firerescuedash.data import charger_donnees
from firerescuedash.nettoyage import nettoyer_noms_colonnes, convertir_colonnes_numeriques
from firerescuedash.export_excel import exporter_excel


CHEMIN_DONNEES = "data/interventions2023.csv"


def main():
    df = charger_donnees(CHEMIN_DONNEES)

    df = nettoyer_noms_colonnes(df)
    df = convertir_colonnes_numeriques(df)

    exporter_excel(df)

    print("\nProjet exécuté avec succès.")


if __name__ == "__main__":
    main()