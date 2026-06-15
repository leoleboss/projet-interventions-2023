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

    Cette fonction orchestre les grandes étapes de l'export.
    Les responsabilités détaillées sont réparties dans des fonctions dédiées.
    """
    os.makedirs("outputs", exist_ok=True)
    chemin_sortie = "outputs/reporting_interventions_2023.xlsx"

    contexte = preparer_contexte(df)

    workbook, feuilles = creer_classeur()

    ecrire_donnees_nettoyees(feuilles["donnees"], df)

    contexte["derniere_ligne_donnees"] = feuilles["donnees"].max_row
    contexte["derniere_ligne_calculs"] = len(contexte["departements"]) + 1

    ecrire_listes_filtres(feuilles["listes"], contexte)
    ecrire_calculs(feuilles["calculs"], contexte)
    construire_dashboard(feuilles["dashboard"], contexte)
    ajouter_graphiques(feuilles["dashboard"])
    appliquer_formats(workbook, feuilles)

    workbook.save(chemin_sortie)

    print("Fichier Excel créé avec formules Excel :", chemin_sortie)


def preparer_contexte(df):
    """
    Prépare les informations réutilisées dans plusieurs feuilles Excel.
    """
    colonnes = list(df.columns)

    colonne_region = trouver_colonne(colonnes, "gion")
    colonne_departement = trouver_colonne(colonnes, "partement")

    colonnes_categories = {
        "Incendies": "incendies",
        "Secours à personne": "secours_à_personne",
        "Accidents": "accidents_de_circulation",
        "Opérations diverses": "opérations_diverses",
    }

    lettres_categories = {}
    for nom_categorie, colonne_df in colonnes_categories.items():
        index_colonne = colonnes.index(colonne_df) + 1
        lettres_categories[nom_categorie] = get_column_letter(index_colonne)

    return {
        "colonnes": colonnes,
        "colonne_region": colonne_region,
        "colonne_departement": colonne_departement,
        "lettre_region": get_column_letter(colonnes.index(colonne_region) + 1),
        "lettre_departement": get_column_letter(colonnes.index(colonne_departement) + 1),
        "colonnes_categories": colonnes_categories,
        "lettres_categories": lettres_categories,
        "regions": ["Toutes"] + sorted(set(df[colonne_region].dropna())),
        "categories": ["Toutes"] + list(colonnes_categories.keys()),
        "departements": sorted(set(df[colonne_departement].dropna())),
    }


def creer_classeur():
    """
    Crée le classeur Excel et ses différentes feuilles.
    """
    workbook = Workbook()

    dashboard = workbook.active
    dashboard.title = "tableau_de_bord"

    feuilles = {
        "dashboard": dashboard,
        "donnees": workbook.create_sheet("donnees_nettoyees"),
        "calculs": workbook.create_sheet("calculs"),
        "listes": workbook.create_sheet("listes_filtres"),
    }

    return workbook, feuilles


def ecrire_donnees_nettoyees(feuille, df):
    """
    Écrit le DataFrame nettoyé dans la feuille de données.
    """
    for col_index, nom_colonne in enumerate(df.columns, start=1):
        feuille.cell(row=1, column=col_index).value = nom_colonne

    for row_index, ligne in enumerate(df.itertuples(index=False), start=2):
        for col_index, valeur in enumerate(ligne, start=1):
            feuille.cell(row=row_index, column=col_index).value = valeur

    feuille.auto_filter.ref = feuille.dimensions


def ecrire_listes_filtres(feuille, contexte):
    """
    Écrit les listes utilisées par les filtres du dashboard.
    """
    feuille["A1"] = "regions"
    for i, region in enumerate(contexte["regions"], start=2):
        feuille[f"A{i}"] = region

    feuille["B1"] = "categories"
    for i, categorie in enumerate(contexte["categories"], start=2):
        feuille[f"B{i}"] = categorie


def formule_sumifs_departement(categorie, ligne, contexte):
    """
    Construit une formule Excel SUMIFS pour une catégorie et un département.
    """
    colonne_lettre = contexte["lettres_categories"][categorie]
    lettre_departement = contexte["lettre_departement"]
    lettre_region = contexte["lettre_region"]
    derniere_ligne = contexte["derniere_ligne_donnees"]

    return (
        f'=SUMIFS(donnees_nettoyees!${colonne_lettre}$2:${colonne_lettre}${derniere_ligne},'
        f'donnees_nettoyees!${lettre_departement}$2:${lettre_departement}${derniere_ligne},A{ligne},'
        f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne},'
        f'IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6))'
    )


def formule_total_filtre(ligne):
    """
    Construit la formule Excel du total filtré par catégorie.
    """
    return (
        f'=IF(tableau_de_bord!$F$6="Toutes",SUM(B{ligne}:E{ligne}),'
        f'IF(tableau_de_bord!$F$6="Incendies",B{ligne},'
        f'IF(tableau_de_bord!$F$6="Secours à personne",C{ligne},'
        f'IF(tableau_de_bord!$F$6="Accidents",D{ligne},'
        f'IF(tableau_de_bord!$F$6="Opérations diverses",E{ligne},0)))))'
    )


def formule_total_categorie(categorie, colonne_lettre, ligne, contexte):
    """
    Construit la formule Excel du total par grande catégorie.
    """
    lettre_region = contexte["lettre_region"]
    derniere_ligne = contexte["derniere_ligne_donnees"]

    return (
        f'=IF(tableau_de_bord!$F$6="Toutes",'
        f'SUMIFS(donnees_nettoyees!${colonne_lettre}$2:${colonne_lettre}${derniere_ligne},'
        f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne},'
        f'IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6)),'
        f'IF(tableau_de_bord!$F$6=H{ligne},'
        f'SUMIFS(donnees_nettoyees!${colonne_lettre}$2:${colonne_lettre}${derniere_ligne},'
        f'donnees_nettoyees!${lettre_region}$2:${lettre_region}${derniere_ligne},'
        f'IF(tableau_de_bord!$C$6="Toutes","*",tableau_de_bord!$C$6)),0))'
    )


def ecrire_calculs(calculs, contexte):
    """
    Crée la feuille de calculs avec les formules Excel.
    """
    ecrire_entetes_calculs(calculs)
    ecrire_calculs_departements(calculs, contexte)
    ecrire_calculs_categories(calculs, contexte)


def ecrire_entetes_calculs(calculs):
    """
    Écrit les en-têtes de la feuille calculs.
    """
    calculs["A1"] = "departement"
    calculs["B1"] = "incendies"
    calculs["C1"] = "secours_personne"
    calculs["D1"] = "accidents"
    calculs["E1"] = "operations_diverses"
    calculs["F1"] = "total_filtre"


def ecrire_calculs_departements(calculs, contexte):
    """
    Écrit les calculs par département.
    """
    for ligne, departement in enumerate(contexte["departements"], start=2):
        calculs[f"A{ligne}"] = departement

        calculs[f"B{ligne}"] = formule_sumifs_departement("Incendies", ligne, contexte)
        calculs[f"C{ligne}"] = formule_sumifs_departement("Secours à personne", ligne, contexte)
        calculs[f"D{ligne}"] = formule_sumifs_departement("Accidents", ligne, contexte)
        calculs[f"E{ligne}"] = formule_sumifs_departement("Opérations diverses", ligne, contexte)
        calculs[f"F{ligne}"] = formule_total_filtre(ligne)


def ecrire_calculs_categories(calculs, contexte):
    """
    Écrit les calculs par grande catégorie.
    """
    calculs["H1"] = "categorie"
    calculs["I1"] = "interventions"

    for ligne, (categorie, colonne_lettre) in enumerate(
        contexte["lettres_categories"].items(),
        start=2,
    ):
        calculs[f"H{ligne}"] = categorie
        calculs[f"I{ligne}"] = formule_total_categorie(
            categorie,
            colonne_lettre,
            ligne,
            contexte,
        )


def construire_dashboard(dashboard, contexte):
    """
    Construit le tableau de bord : titre, filtres, tableaux et KPI.
    """
    styles = creer_styles_dashboard()

    largeur_colonnes(dashboard)

    creer_titre_dashboard(dashboard, styles)
    creer_filtres_dashboard(dashboard, contexte, styles)
    creer_table_top_departements(dashboard, contexte, styles)
    creer_table_categories(dashboard, styles)
    creer_kpis(dashboard, styles)


def creer_styles_dashboard():
    """
    Centralise les styles utilisés dans le dashboard.
    """
    rouge = "A61C1C"
    rouge_clair = "FCE5E5"
    gris_clair = "F3F3F3"

    return {
        "remplissage_rouge": PatternFill("solid", fgColor=rouge),
        "remplissage_rouge_clair": PatternFill("solid", fgColor=rouge_clair),
        "remplissage_gris": PatternFill("solid", fgColor=gris_clair),
        "titre_font": Font(size=18, bold=True, color="FFFFFF"),
        "sous_titre_font": Font(size=11, bold=True, color="FFFFFF"),
        "kpi_font": Font(size=12, bold=True),
        "valeur_kpi_font": Font(size=15, bold=True),
        "bordure": Border(
            left=Side(style="thin", color=rouge),
            right=Side(style="thin", color=rouge),
            top=Side(style="thin", color=rouge),
            bottom=Side(style="thin", color=rouge),
        ),
    }


def creer_titre_dashboard(dashboard, styles):
    """
    Crée le titre principal et la zone des filtres.
    """
    dashboard.merge_cells("A1:N2")
    dashboard["A1"] = "TABLEAU DE BORD - INTERVENTIONS DES SERVICES DE SECOURS 2023"
    dashboard["A1"].fill = styles["remplissage_rouge"]
    dashboard["A1"].font = styles["titre_font"]
    dashboard["A1"].alignment = Alignment(horizontal="center", vertical="center")

    dashboard.merge_cells("A4:N5")
    dashboard["A4"] = "Zone de filtres"
    dashboard["A4"].fill = styles["remplissage_rouge_clair"]
    dashboard["A4"].alignment = Alignment(horizontal="center", vertical="center")


def creer_filtres_dashboard(dashboard, contexte, styles):
    """
    Crée les cellules de filtres et les listes déroulantes.
    """
    dashboard["B6"] = "Filtre région"
    dashboard["C6"] = "Toutes"

    dashboard["E6"] = "Filtre catégorie"
    dashboard["F6"] = "Toutes"

    for cellule in ["B6", "E6"]:
        dashboard[cellule].fill = styles["remplissage_rouge"]
        dashboard[cellule].font = Font(bold=True, color="FFFFFF")
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = styles["bordure"]

    for cellule in ["C6", "F6"]:
        dashboard[cellule].fill = styles["remplissage_gris"]
        dashboard[cellule].alignment = Alignment(horizontal="center")
        dashboard[cellule].border = styles["bordure"]

    validation_region = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$A$2:$A${len(contexte['regions']) + 1}",
        allow_blank=False,
    )

    validation_categorie = DataValidation(
        type="list",
        formula1=f"'listes_filtres'!$B$2:$B${len(contexte['categories']) + 1}",
        allow_blank=False,
    )

    dashboard.add_data_validation(validation_region)
    dashboard.add_data_validation(validation_categorie)

    validation_region.add(dashboard["C6"])
    validation_categorie.add(dashboard["F6"])


def creer_table_top_departements(dashboard, contexte, styles):
    """
    Crée le tableau du top 10 des départements.
    """
    derniere_ligne_calculs = contexte["derniere_ligne_calculs"]

    dashboard["A9"] = "Top 10 départements"
    dashboard["A9"].fill = styles["remplissage_rouge"]
    dashboard["A9"].font = styles["sous_titre_font"]

    dashboard["A10"] = "Département"
    dashboard["B10"] = "Interventions"

    for cellule in ["A10", "B10"]:
        dashboard[cellule].fill = styles["remplissage_gris"]
        dashboard[cellule].font = Font(bold=True)
        dashboard[cellule].border = styles["bordure"]

    for i in range(1, 11):
        ligne_dashboard = 10 + i

        dashboard[f"B{ligne_dashboard}"] = (
            f'=LARGE(calculs!$F$2:$F${derniere_ligne_calculs},{i})'
        )

        dashboard[f"A{ligne_dashboard}"] = (
            f'=INDEX(calculs!$A$2:$A${derniere_ligne_calculs},'
            f'MATCH(B{ligne_dashboard},calculs!$F$2:$F${derniere_ligne_calculs},0))'
        )


def creer_table_categories(dashboard, styles):
    """
    Crée le tableau des grandes catégories.
    """
    dashboard["D9"] = "Grandes catégories"
    dashboard["D9"].fill = styles["remplissage_rouge"]
    dashboard["D9"].font = styles["sous_titre_font"]

    dashboard["D10"] = "Catégorie"
    dashboard["E10"] = "Interventions"

    for cellule in ["D10", "E10"]:
        dashboard[cellule].fill = styles["remplissage_gris"]
        dashboard[cellule].font = Font(bold=True)
        dashboard[cellule].border = styles["bordure"]

    for ligne in range(2, 6):
        ligne_dashboard = ligne + 9
        dashboard[f"D{ligne_dashboard}"] = f"=calculs!H{ligne}"
        dashboard[f"E{ligne_dashboard}"] = f"=calculs!I{ligne}"


def creer_kpis(dashboard, styles):
    """
    Crée les indicateurs KPI du dashboard.
    """
    kpis = {
        "J9": ("Total interventions", "=SUM(E11:E14)"),
        "L9": ("Département n°1", "=A11"),
        "J13": ("Part secours à personne", '=IF(J10=0,0,E12/J10)'),
        "L13": (
            "Moyenne / département",
            '=IF(COUNTIF(B11:B20,">0")=0,0,J10/COUNTIF(B11:B20,">0"))',
        ),
    }

    for position, (nom, formule) in kpis.items():
        dashboard[position] = nom
        dashboard[position].fill = styles["remplissage_rouge_clair"]
        dashboard[position].font = styles["kpi_font"]
        dashboard[position].alignment = Alignment(horizontal="center")
        dashboard[position].border = styles["bordure"]

        cellule_valeur = dashboard.cell(
            row=dashboard[position].row + 1,
            column=dashboard[position].column,
        )
        cellule_valeur.value = formule
        cellule_valeur.font = styles["valeur_kpi_font"]
        cellule_valeur.alignment = Alignment(horizontal="center")
        cellule_valeur.border = styles["bordure"]


def ajouter_graphiques(dashboard):
    """
    Ajoute les graphiques Excel au dashboard.
    """
    ajouter_graphique_top_departements(dashboard)
    ajouter_graphique_categories(dashboard)



def ajouter_graphique_top_departements(dashboard):     
    graphique_top = BarChart()
    graphique_top.type = "bar"
    graphique_top.style = 10
    graphique_top.title = "Top 10 départements"
    graphique_top.legend = None
    graphique_top.height = 8
    graphique_top.width = 18

    # Afficher les axes
    graphique_top.x_axis.delete = False
    graphique_top.y_axis.delete = False

    # Ne pas afficher de titres d'axes
    graphique_top.x_axis.title = None
    graphique_top.y_axis.title = None

    data_top = Reference(dashboard, min_col=2, min_row=10, max_row=20)
    labels_top = Reference(dashboard, min_col=1, min_row=11, max_row=20)

    graphique_top.add_data(data_top, titles_from_data=True)
    graphique_top.set_categories(labels_top)

    # Étiquettes de données : afficher uniquement les valeurs
    graphique_top.dLbls = DataLabelList()
    graphique_top.dLbls.showVal = True
    graphique_top.dLbls.showCatName = False
    graphique_top.dLbls.showSerName = False
    graphique_top.dLbls.showPercent = False
    graphique_top.dLbls.showLegendKey = False
    graphique_top.dLbls.numFmt = "#,##0"

    dashboard.add_chart(graphique_top, "A25")


def ajouter_graphique_categories(dashboard):
    """
    Ajoute le graphique de répartition des interventions.
    """
    graphique_categories = PieChart()
    graphique_categories.title = "Répartition des interventions"
    graphique_categories.height = 8
    graphique_categories.width = 14

    data_categories = Reference(
        dashboard,
        min_col=5,
        min_row=10,
        max_row=14
    )

    labels_categories = Reference(
        dashboard,
        min_col=4,
        min_row=11,
        max_row=14
    )

    graphique_categories.add_data(
        data_categories,
        titles_from_data=True
    )

    graphique_categories.set_categories(labels_categories)

    # Afficher uniquement le %
    graphique_categories.dLbls = DataLabelList()
    graphique_categories.dLbls.showPercent = True
    graphique_categories.dLbls.showVal = False
    graphique_categories.dLbls.showCatName = False
    graphique_categories.dLbls.showSerName = False
    graphique_categories.dLbls.showLegendKey = False

    graphique_categories.legend.position = "b"

    dashboard.add_chart(graphique_categories, "I25")


def appliquer_formats(workbook, feuilles):
    """
    Applique les formats numériques et options finales du classeur.
    """
    dashboard = feuilles["dashboard"]
    donnees = feuilles["donnees"]
    calculs = feuilles["calculs"]
    listes = feuilles["listes"]

    largeur_colonnes(calculs)
    largeur_colonnes(donnees)

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
