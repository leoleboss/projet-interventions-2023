from firerescuedash.data import charger_donnees
from firerescuedash.nettoyage import (
    nettoyer_noms_colonnes,
    convertir_colonnes_numeriques
)
from firerescuedash.export_excel import exporter_excel


def main():

    # Chargement des données depuis MinIO
    df = charger_donnees()

    # Nettoyage
    df = nettoyer_noms_colonnes(df)
    df = convertir_colonnes_numeriques(df)

    # Génération du fichier Excel
    exporter_excel(df)

    print("\nProjet exécuté avec succès.")


if __name__ == "__main__":
    main()
    