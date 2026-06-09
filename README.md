# FireRescueDash

Automatisation d'un reporting Excel sur les interventions des services de secours en 2023.

## Contexte

Ce projet a pour objectif de créer automatiquement un tableau de bord Excel à partir d'un fichier de données.

Le tableau de bord permet d'analyser les interventions réalisées par les services de secours en France en 2023.

L'utilisateur peut consulter les principaux indicateurs, filtrer les données et visualiser les départements les plus concernés par les interventions.

---

## Données

### Source

Les données proviennent du site data.gouv.fr :

https://www.data.gouv.fr/datasets/interventions-realisees-par-les-services-d-incendie-et-de-secours

Le fichier utilisé contient les interventions réalisées par les services de secours en France en 2023.

Le jeu de données contient 99 lignes et 71 colonnes.

### Principales variables utilisées

| Variable                 | Description                                        |
| ------------------------ | -------------------------------------------------- |
| Année                    | Année d'observation                                |
| Zone                     | Zone géographique                                  |
| Région                   | Région administrative                              |
| Numéro                   | Numéro du département                              |
| Département              | Nom du département                                 |
| Incendies                | Nombre d'interventions liées aux incendies         |
| Secours à personne       | Nombre d'interventions liées au secours à personne |
| Accidents de circulation | Nombre d'interventions liées aux accidents         |
| Opérations diverses      | Nombre d'opérations diverses                       |
| Total interventions      | Nombre total d'interventions                       |

---

## Objectif du projet

L'objectif est de produire automatiquement un tableau de bord Excel à partir des données brutes.

Le script Python permet :

* de charger les données ;
* de nettoyer les colonnes ;
* de convertir les valeurs numériques ;
* de générer automatiquement un fichier Excel ;
* de créer un tableau de bord avec filtres, KPI et graphiques.

Les calculs du tableau de bord sont réalisés directement dans Excel à l'aide de formules.

---

## Schéma du projet

```text
Données brutes
      ↓
Chargement des données
      ↓
Nettoyage des données
      ↓
Création du fichier Excel
      ↓
Formules Excel
      ↓
Graphiques Excel
      ↓
Tableau de bord final
```

---

## Aperçu du résultat

Le fichier Excel généré contient :

* une feuille de données nettoyées ;
* une feuille de calculs ;
* un tableau de bord ;
* des KPI ;
* des filtres ;
* des graphiques dynamiques.

Les calculs restent visibles dans Excel afin que l'utilisateur puisse comprendre et réutiliser le fichier.

---

## Fonctionnement du tableau de bord

Le tableau de bord utilise :

* des listes déroulantes ;
* des formules Excel ;
* des graphiques Excel ;
* des indicateurs calculés automatiquement.

Les principales fonctions utilisées sont :

* SUMIFS ;
* IF ;
* INDEX ;
* MATCH ;
* LARGE ;
* COUNTIF.

Les graphiques sont alimentés directement par les cellules de calcul Excel.

---

## Installation

Cloner le projet :

```bash
git clone <url_du_projet>
cd projet-interventions-2023
```

Installer les dépendances :

```bash
uv sync
```

---

## Utilisation

Lancer le projet :

```bash
uv run firerescuedash
```

Le fichier Excel est généré automatiquement dans le dossier de sortie.

---

## Structure du projet

```text
projet-interventions-2023/
├── pyproject.toml
├── uv.lock
├── README.md
│
├── notebooks/
│   └── 01_exploration.ipynb
│
└── src/
    └── firerescuedash/
        ├── __init__.py
        ├── analyse.py
        ├── data.py
        ├── export_excel.py
        ├── main.py
        └── nettoyage.py
```

---

## Rôle de chaque fichier

### data.py

Charge les données.

### nettoyage.py

Nettoie les noms de colonnes et convertit les colonnes numériques.

### analyse.py

Contient les fonctions d'analyse utilisées pendant l'exploration des données.

### export_excel.py

Construit le fichier Excel :

* feuilles ;
* formules ;
* KPI ;
* graphiques ;
* tableau de bord.

### main.py

Lance le projet complet.

---

## Bibliothèques utilisées

* pandas
* openpyxl
* uv
* openpyxl.chart

```
```

