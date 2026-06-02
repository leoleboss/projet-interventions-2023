import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


def exporter_excel(df, top_departements, categories):
    chemin_sortie = "outputs/reporting_interventions_2023.xlsx"

    colonne_region = [col for col in df.columns if "gion" in col][0]
    colonne_departement = [col for col in df.columns if "partement" in col][0]

    zones_uniques = ["Toutes"] + sorted(df["zone"].dropna().unique())
    regions_uniques = ["Toutes"] + sorted(df[colonne_region].dropna().unique())
    departements_uniques = sorted(df[colonne_departement].dropna().unique())

    filtres_df = pd.DataFrame({
        "zones": pd.Series(zones_uniques),
        "regions": pd.Series(regions_uniques)
    })

    calcul_df = pd.DataFrame({
        "departement": departements_uniques
    })

    with pd.ExcelWriter(chemin_sortie, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="donnees_nettoyees", index=False)
        filtres_df.to_excel(writer, sheet_name="listes_filtres", index=False)
        calcul_df.to_excel(writer, sheet_name="calcul_dashboard", index=False)

    workbook = load_workbook(chemin_sortie)

    dashboard = workbook.create_sheet("tableau_de_bord", 0)
    feuille_filtres = workbook["listes_filtres"]
    feuille_calcul = workbook["calcul_dashboard"]
    feuille_donnees = workbook["donnees_nettoyees"]

    vert = "6AA84F"
    vert_clair = "D9EAD3"
    gris_clair = "F3F3F3"

    titre_font = Font(size=18, bold=True, color="FFFFFF")
    sous_titre_font = Font(size=11, bold=True, color="FFFFFF")
    kpi_font = Font(size=12, bold=True)
    valeur_kpi_font = Font(size=15, bold=True)

    remplissage_vert = PatternFill("solid", fgColor=vert)
    remplissage_vert_clair = PatternFill("solid", fgColor=vert_clair)
    remplissage_gris = PatternFill("solid", fgColor=gris_clair)

    bordure = Border(
        left=Side(style="thin", color=vert),
        right=Side(style="thin", color=vert),
        top=Side(style="thin", color=vert),
        bottom=Side(style="thin", color=vert)
    )

    for col in range(1, 15):
        dashboard.column_dimensions[chr(64 + col)].width = 16

    dashboard.row_dimensions[1].height = 28
    dashboard.row_dimensions[2].height = 28
    dashboard.row_dimensions[4].height = 25
    dashboard.row_dimensions[5].height = 25
    dashboard.row_dimensions[6].height = 25
    dashboard.row_dimensions[7].height = 18
    dashboard.row_dimensions[8].height = 18
    dashboard.row_dimensions[21].height = 25

    dashboard.merge_cells("A1:N2")
    dashboard["A1"] = "Tableau de bord - Interventions des secours en 2023"
    dashboard["A1"].fill = remplissage_vert
    dashboard["A1"].font = titre_font
    dashboard["A1"].alignment = Alignment(horizontal="center", vertical="center")

    dashboard.merge_cells("A4:N5")
    dashboard["A4"] = "Zone de filtres"
    dashboard["A4"].fill = remplissage_vert_clair
    dashboard["A4"].alignment = Alignment(horizontal="center", vertical="center")
    dashboard["A4"].border = bordure

    dashboard["B6"] = "Filtre zone"
    dashboard["C6"] = "Toutes"
    dashboard["E6"] = "Filtre région"
    dashboard["F6"] = "Toutes"

    for cellule in ["B6", "E6"]:
        dashboard[cellule].fill = remplissage_vert
        dashboard[cellule].font = Font(bold=True, color="FFFFFF")
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = bordure

    for cellule in ["C6", "F6"]:
        dashboard[cellule].fill = remplissage_gris
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = bordure

    validation_zone = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$A$2:$A${feuille_filtres.max_row}",
        allow_blank=False
    )

    validation_region = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$B$2:$B${feuille_filtres.max_row}",
        allow_blank=False
    )

    dashboard.add_data_validation(validation_zone)
    dashboard.add_data_validation(validation_region)

    validation_zone.add(dashboard["C6"])
    validation_region.add(dashboard["F6"])

    critere_zone = 'IF($C$6="Toutes","*",$C$6)'
    critere_region = 'IF($F$6="Toutes","*",$F$6)'

    dashboard["A9"] = "Top 10 départements"
    dashboard["A9"].fill = remplissage_vert
    dashboard["A9"].font = sous_titre_font

    dashboard["A10"] = "Département"
    dashboard["B10"] = "Interventions"

    for cell in ["A10", "B10"]:
        dashboard[cell].fill = remplissage_gris
        dashboard[cell].font = Font(bold=True)
        dashboard[cell].border = bordure

    feuille_calcul["B1"] = "interventions_filtrees"

    for ligne in range(2, feuille_calcul.max_row + 1):
        feuille_calcul[f"B{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!$BS$2:$BS$200,'
            f'donnees_nettoyees!$E$2:$E$200,A{ligne},'
            f'donnees_nettoyees!$B$2:$B$200,tableau_de_bord!$C$6,'
            f'donnees_nettoyees!$C$2:$C$200,tableau_de_bord!$F$6)'
        )

        feuille_calcul[f"C{ligne}"] = (
            f'=SUMIFS(donnees_nettoyees!$BS$2:$BS$200,'
            f'donnees_nettoyees!$E$2:$E$200,A{ligne},'
            f'donnees_nettoyees!$B$2:$B$200,IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6),'
            f'donnees_nettoyees!$C$2:$C$200,IF(tableau_de_bord!$F$6="Toutes","*",tableau_de_bord!$F$6))'
        )

    for i in range(1, 11):
        ligne_dashboard = 10 + i
        dashboard[f"B{ligne_dashboard}"] = f'=LARGE(calcul_dashboard!$C$2:$C$200,{i})'
        dashboard[f"A{ligne_dashboard}"] = (
            f'=INDEX(calcul_dashboard!$A$2:$A$200,'
            f'MATCH(B{ligne_dashboard},calcul_dashboard!$C$2:$C$200,0))'
        )

    dashboard["D9"] = "Grandes catégories"
    dashboard["D9"].fill = remplissage_vert
    dashboard["D9"].font = sous_titre_font

    dashboard["D10"] = "Catégorie"
    dashboard["E10"] = "Interventions"

    for cell in ["D10", "E10"]:
        dashboard[cell].fill = remplissage_gris
        dashboard[cell].font = Font(bold=True)
        dashboard[cell].border = bordure

    categories_lignes = [
        ("Incendies", "R"),
        ("Secours à personne", "AM"),
        ("Accidents", "AS"),
        ("Opérations diverses", "BR")
    ]

    for index, (nom, colonne_excel) in enumerate(categories_lignes, start=11):
        dashboard.cell(row=index, column=4).value = nom
        dashboard.cell(row=index, column=5).value = (
            f'=SUMIFS(donnees_nettoyees!${colonne_excel}$2:${colonne_excel}$200,'
            f'donnees_nettoyees!$B$2:$B$200,{critere_zone},'
            f'donnees_nettoyees!$C$2:$C$200,{critere_region})'
        )

    kpis = {
        "J9": ("Total interventions", f'=SUM(E11:E14)'),
        "L9": ("Nombre de lignes", f'=COUNTIFS(donnees_nettoyees!$B$2:$B$200,{critere_zone},donnees_nettoyees!$C$2:$C$200,{critere_region})'),
        "J13": ("Part secours à personne", "=E12/J10"),
        "L13": ("Nombre de catégories", 4)
    }

    for position, (nom, valeur) in kpis.items():
        dashboard[position] = nom
        dashboard[position].fill = remplissage_vert_clair
        dashboard[position].font = kpi_font
        dashboard[position].alignment = Alignment(horizontal="center")
        dashboard[position].border = bordure

        cellule_valeur = dashboard.cell(
            row=dashboard[position].row + 1,
            column=dashboard[position].column
        )
        cellule_valeur.value = valeur
        cellule_valeur.font = valeur_kpi_font
        cellule_valeur.alignment = Alignment(horizontal="center")
        cellule_valeur.border = bordure

    dashboard["J14"].number_format = "0.0%"

        # Graphique top départements
    graphique_top = BarChart()
    graphique_top.type = "bar"
    graphique_top.style = 10
    graphique_top.title = "Top 10 départements"
    graphique_top.y_axis.title = ""
    graphique_top.x_axis.title = ""
    graphique_top.legend = None

    graphique_top.height = 8
    graphique_top.width = 18

    data_top = Reference(dashboard, min_col=2, min_row=10, max_row=20)
    labels_top = Reference(dashboard, min_col=1, min_row=11, max_row=20)

    graphique_top.add_data(data_top, titles_from_data=True)
    graphique_top.set_categories(labels_top)

    graphique_top.dLbls = DataLabelList()
    graphique_top.dLbls.showVal = True
    graphique_top.dLbls.showCatName = True
    graphique_top.dLbls.showSerName = False
    graphique_top.dLbls.showLegendKey = False

    dashboard.add_chart(graphique_top, "A25")

        # Graphique catégories
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
    graphique_categories.dLbls.showVal = False
    graphique_categories.dLbls.showCatName = False
    graphique_categories.dLbls.showSerName = False
    graphique_categories.dLbls.showLegendKey = False

    # Légende en bas
    graphique_categories.legend.position = "b"

    dashboard.add_chart(graphique_categories, "I25")

    feuille_donnees.auto_filter.ref = feuille_donnees.dimensions

    feuille_filtres.sheet_state = "hidden"
    feuille_calcul.sheet_state = "hidden"

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    workbook.save(chemin_sortie)

    print("Fichier Excel créé avec tableau de bord dynamique :", chemin_sortie)



