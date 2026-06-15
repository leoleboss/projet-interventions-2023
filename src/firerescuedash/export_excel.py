import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


def trouver_colonne(colonnes, mot):
    """
    Trouve une colonne à partir d'un morceau de son nom.
    """
    for colonne in colonnes:
        if mot in colonne:
            return colonne
    raise ValueError(f"Colonne contenant '{mot}' introuvable.")


def largeur_colonnes(feuille, largeur=16):
    """
    Applique une largeur simple aux colonnes.
    """
    for col in range(1, 15):
        feuille.column_dimensions[get_column_letter(col)].width = largeur


def exporter_excel(df):
    """
    Crée un fichier Excel avec :
    - une feuille de données nettoyées ;
    - une feuille de calculs avec des formules Excel ;
    - un tableau de bord avec filtres, KPI et graphiques.
    """

    os.makedirs("outputs", exist_ok=True)
    chemin_sortie = "outputs/reporting_interventions_2023.xlsx"

    # Colonnes importantes
    colonnes = list(df.columns)

    colonne_region = trouver_colonne(colonnes, "gion")
    colonne_departement = trouver_colonne(colonnes, "partement")

    # Colonnes utilisées pour les calculs
    colonnes_categories = {
    "Incendies": "incendies",
    "Secours à personne": "secours_à_personne",
    "Accidents": "accidents_de_circulation",
    "Opérations diverses": "opérations_diverses",
}

    # Création du classeur
    workbook = Workbook()

    # Feuilles
    dashboard = workbook.active
    dashboard.title = "tableau_de_bord"

    donnees = workbook.create_sheet("donnees_nettoyees")
    calculs = workbook.create_sheet("calculs")
    listes = workbook.create_sheet("listes_filtres")

    # ------------------------------------------------------------------
    # 1. Écriture des données nettoyées
    # ------------------------------------------------------------------

    for col_index, nom_colonne in enumerate(df.columns, start=1):
        donnees.cell(row=1, column=col_index).value = nom_colonne

    for row_index, ligne in enumerate(df.itertuples(index=False), start=2):
        for col_index, valeur in enumerate(ligne, start=1):
            donnees.cell(row=row_index, column=col_index).value = valeur

    donnees.auto_filter.ref = donnees.dimensions

    derniere_ligne_donnees = donnees.max_row

    # Récupérer les lettres Excel des colonnes utiles
    index_region = colonnes.index(colonne_region) + 1
    index_departement = colonnes.index(colonne_departement) + 1

    lettre_region = get_column_letter(index_region)
    lettre_departement = get_column_letter(index_departement)

    lettres_categories = {}
    for nom_categorie, colonne_df in colonnes_categories.items():
        index_colonne = colonnes.index(colonne_df) + 1
        lettres_categories[nom_categorie] = get_column_letter(index_colonne)

    # ------------------------------------------------------------------
    # 2. Listes de filtres
    # ------------------------------------------------------------------

    regions = ["Toutes"] + sorted(set(df[colonne_region].dropna()))
    categories = ["Toutes"] + list(colonnes_categories.keys())

    listes["A1"] = "regions"
    for i, region in enumerate(regions, start=2):
        listes[f"A{i}"] = region

    listes["B1"] = "categories"
    for i, categorie in enumerate(categories, start=2):
        listes[f"B{i}"] = categorie

    # ------------------------------------------------------------------
    # 3. Feuille calculs avec formules Excel
    # ------------------------------------------------------------------

    calculs["A1"] = "departement"
    calculs["B1"] = "incendies"
    calculs["C1"] = "secours_personne"
    calculs["D1"] = "accidents"
    calculs["E1"] = "operations_diverses"
    calculs["F1"] = "total_filtre"

    departements = sorted(set(df[colonne_departement].dropna()))

    for ligne, departement in enumerate(departements, start=2):
        calculs[f"A{ligne}"] = departement

        calculs[f"B{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!${lettres_categories["Incendies"]}$2:${lettres_categories["Incendies"]}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_departement}$2:${lettre_departement}${derniere_ligne_donnees},A{ligne},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6))'
        )

        calculs[f"C{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!${lettres_categories["Secours à personne"]}$2:${lettres_categories["Secours à personne"]}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_departement}$2:${lettre_departement}${derniere_ligne_donnees},A{ligne},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6))'
        )

        calculs[f"D{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!${lettres_categories["Accidents"]}$2:${lettres_categories["Accidents"]}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_departement}$2:${lettre_departement}${derniere_ligne_donnees},A{ligne},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6))'
        )

        calculs[f"E{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!${lettres_categories["Opérations diverses"]}$2:${lettres_categories["Opérations diverses"]}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_departement}$2:${lettre_departement}${derniere_ligne_donnees},A{ligne},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6))'
        )

        calculs[f"F{ligne}"] = (
            f'=IF(tableau_de_bord!$F$6="Toutes",SUM(B{ligne}:E{ligne}),'
            f'IF(tableau_de_bord!$F$6="Incendies",B{ligne},'
            f'IF(tableau_de_bord!$F$6="Secours à personne",C{ligne},'
            f'IF(tableau_de_bord!$F$6="Accidents",D{ligne},'
            f'IF(tableau_de_bord!$F$6="Opérations diverses",E{ligne},0)))))'
        )

    derniere_ligne_calculs = len(departements) + 1

    # Calculs des grandes catégories
    calculs["H1"] = "categorie"
    calculs["I1"] = "interventions"

    for ligne, (categorie, colonne_lettre) in enumerate(lettres_categories.items(), start=2):
        calculs[f"H{ligne}"] = categorie
        calculs[f"I{ligne}"] = (
            f'=IF(tableau_de_bord!$F$6="Toutes",'
            f'SUMIFS(donnees_nettoyees!${colonne_lettre}$2:${colonne_lettre}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},'
            f'IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6)),'
            f'IF(tableau_de_bord!$F$6=H{ligne},'
            f'SUMIFS(donnees_nettoyees!${colonne_lettre}$2:${colonne_lettre}${derniere_ligne_donnees},'
            f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne_donnees},'
            f'IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6)),0))'
        )

    # ------------------------------------------------------------------
    # 4. Mise en forme du tableau de bord
    # ------------------------------------------------------------------

    rouge = "A61C1C"
    rouge_clair = "FCE5E5"
    gris_clair = "F3F3F3"

    remplissage_rouge = PatternFill("solid", fgColor=rouge)
    remplissage_rouge_clair = PatternFill("solid", fgColor=rouge_clair)
    remplissage_gris = PatternFill("solid", fgColor=gris_clair)

    titre_font = Font(size=18, bold=True, color="FFFFFF")
    sous_titre_font = Font(size=11, bold=True, color="FFFFFF")
    kpi_font = Font(size=12, bold=True)
    valeur_kpi_font = Font(size=15, bold=True)

    bordure = Border(
        left=Side(style="thin", color=rouge),
        right=Side(style="thin", color=rouge),
        top=Side(style="thin", color=rouge),
        bottom=Side(style="thin", color=rouge),
    )

    largeur_colonnes(dashboard)
    largeur_colonnes(calculs)
    largeur_colonnes(donnees)

    dashboard.merge_cells("A1:N2")
    dashboard["A1"] = "TABLEAU DE BORD - INTERVENTIONS DES SERVICES DE SECOURS 2023"
    dashboard["A1"].fill = remplissage_rouge
    dashboard["A1"].font = titre_font
    dashboard["A1"].alignment = Alignment(horizontal="center", vertical="center")

    dashboard.merge_cells("A4:N5")
    dashboard["A4"] = "Zone de filtres"
    dashboard["A4"].fill = remplissage_rouge_clair
    dashboard["A4"].alignment = Alignment(horizontal="center", vertical="center")

    dashboard["B6"] = "Filtre région"
    dashboard["C6"] = "Toutes"

    dashboard["E6"] = "Filtre catégorie"
    dashboard["F6"] = "Toutes"

    for cellule in ["B6", "E6"]:
        dashboard[cellule].fill = remplissage_rouge
        dashboard[cellule].font = Font(bold=True, color="FFFFFF")
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = bordure

    for cellule in ["C6", "F6"]:
        dashboard[cellule].fill = remplissage_gris
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = bordure

    # Listes déroulantes
    validation_region = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$A$2:$A${len(regions) + 1}",
        allow_blank=False,
    )

    validation_categorie = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$B$2:$B${len(categories) + 1}",
        allow_blank=False,
    )

    dashboard.add_data_validation(validation_region)
    dashboard.add_data_validation(validation_categorie)

    validation_region.add(dashboard["C6"])
    validation_categorie.add(dashboard["F6"])

    # ------------------------------------------------------------------
    # 5. Tableaux visibles du dashboard avec formules Excel
    # ------------------------------------------------------------------

    dashboard["A9"] = "Top 10 départements"
    dashboard["A9"].fill = remplissage_rouge
    dashboard["A9"].font = sous_titre_font

    dashboard["A10"] = "Département"
    dashboard["B10"] = "Interventions"

    for cellule in ["A10", "B10"]:
        dashboard[cellule].fill = remplissage_gris
        dashboard[cellule].font = Font(bold=True)
        dashboard[cellule].border = bordure

    for i in range(1, 11):
        ligne_dashboard = 10 + i

        dashboard[f"B{ligne_dashboard}"] = (
            f'=LARGE(calculs!$F$2:$F${derniere_ligne_calculs},{i})'
        )

        dashboard[f"A{ligne_dashboard}"] = (
            f'=INDEX(calculs!$A$2:$A${derniere_ligne_calculs},'
            f'MATCH(B{ligne_dashboard},calculs!$F$2:$F${derniere_ligne_calculs},0))'
        )

    dashboard["D9"] = "Grandes catégories"
    dashboard["D9"].fill = remplissage_rouge
    dashboard["D9"].font = sous_titre_font

    dashboard["D10"] = "Catégorie"
    dashboard["E10"] = "Interventions"

    for cellule in ["D10", "E10"]:
        dashboard[cellule].fill = remplissage_gris
        dashboard[cellule].font = Font(bold=True)
        dashboard[cellule].border = bordure

    for ligne in range(2, 6):
        ligne_dashboard = ligne + 9
        dashboard[f"D{ligne_dashboard}"] = f"=calculs!H{ligne}"
        dashboard[f"E{ligne_dashboard}"] = f"=calculs!I{ligne}"

    # KPI
    kpis = {
        "J9": ("Total interventions", "=SUM(E11:E14)"),
        "L9": ("Département n°1", "=A11"),
        "J13": ("Part secours à personne", '=IF(J10=0,0,E12/J10)'),
        "L13": ("Moyenne / département", '=IF(COUNTIF(B11:B20,">0")=0,0,J10/COUNTIF(B11:B20,">0"))'),
    }

    for position, (nom, formule) in kpis.items():
        dashboard[position] = nom
        dashboard[position].fill = remplissage_rouge_clair
        dashboard[position].font = kpi_font
        dashboard[position].alignment = Alignment(horizontal="center")
        dashboard[position].border = bordure

        cellule_valeur = dashboard.cell(
            row=dashboard[position].row + 1,
            column=dashboard[position].column,
        )
        cellule_valeur.value = formule
        cellule_valeur.font = valeur_kpi_font
        cellule_valeur.alignment = Alignment(horizontal="center")
        cellule_valeur.border = bordure

    # ------------------------------------------------------------------
    # 6. Graphiques Excel basés sur les formules
    # ------------------------------------------------------------------

    graphique_top = BarChart()
    graphique_top.type = "bar"
    graphique_top.title = "Top 10 départements"
    graphique_top.legend = None
    graphique_top.height = 8
    graphique_top.width = 18

    data_top = Reference(dashboard, min_col=2, min_row=10, max_row=20)
    labels_top = Reference(dashboard, min_col=1, min_row=11, max_row=20)

    graphique_top.add_data(data_top, titles_from_data=True)
    graphique_top.set_categories(labels_top)

    graphique_top.dLbls = DataLabelList()
    graphique_top.dLbls.showVal = True
    graphique_top.dLbls.numFmt = "#,##0"

    dashboard.add_chart(graphique_top, "A25")

    graphique_categories = PieChart()
    graphique_categories.title = "Répartition des interventions"
    graphique_categories.height = 8
    graphique_categories.width = 14

    data_categories = Reference(dashboard, min_col=5, min_row=10, max_row=14)
    labels_categories = Reference(dashboard, min_col=4, min_row=11, max_row=14)

    graphique_categories.add_data(data_categories, titles_from_data=True)
    graphique_categories.set_categories(labels_categories)

    graphique_categories.dLbls = DataLabelList()
    graphique_categories.dLbls.showPercent = True
    graphique_categories.legend.position = "b"

    dashboard.add_chart(graphique_categories, "I25")

    # ------------------------------------------------------------------
    # 7. Formats et options Excel
    # ------------------------------------------------------------------

    format_nombre = "#,##0"

    for ligne in range(11, 21):
        dashboard[f"B{ligne}"].number_format = format_nombre

    for ligne in range(11, 15):
        dashboard[f"E{ligne}"].number_format = format_nombre

    dashboard["J10"].number_format = format_nombre
    dashboard["J14"].number_format = "0.0%"
    dashboard["L14"].number_format = format_nombre

    for feuille in [donnees, calculs]:
        for ligne in feuille.iter_rows():
            for cellule in ligne:
                if isinstance(cellule.value, (int, float)):
                    cellule.number_format = format_nombre

    listes.sheet_state = "hidden"

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    workbook.save(chemin_sortie)

    print("Fichier Excel créé avec formules Excel :", chemin_sortie)
    