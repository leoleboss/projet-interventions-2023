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

Le fichier utilisé est :

```text
interventions2023.csv
```

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

Le script Python permet de :

* charger les données ;
* nettoyer les colonnes ;
* convertir les valeurs numériques ;
* réaliser plusieurs analyses ;
* calculer des indicateurs ;
* créer des filtres ;
* générer un tableau de bord Excel dynamique.

---

## Aperçu du résultat

Le fichier Excel généré contient :

* un tableau de bord principal ;
* des indicateurs KPI ;
* des filtres par région et catégorie ;
* un classement des départements ;
* des graphiques dynamiques ;
* les données nettoyées utilisées pour les calculs.

Le fichier final généré est :

```text
outputs/reporting_interventions_2023.xlsx
```

---

## Installation

Cloner le projet :

```bash
git clone <url_du_projet>
cd projet-interventions-2023
```

Installer les dépendances :

```bash
pip install -r requirements.txt
```

---

## Utilisation

Lancer le projet :

```bash
python src/main.py
```

Le fichier Excel est automatiquement créé dans :

```text
outputs/reporting_interventions_2023.xlsx
```

---

## Structure du projet

```text
projet-interventions-2023/
├── data/
│   └── interventions2023.csv
│
├── outputs/
│   └── reporting_interventions_2023.xlsx
│
├── notebooks/
│   └── 01_exploration.ipynb
│
├── src/
│   ├── analyse.py
│   ├── chargement.py
│   ├── export_excel.py
│   ├── main.py
│   └── nettoyage.py
│
├── README.md
├── requirements.txt
└── .gitignore
```

---

## Rôle de chaque fichier

### chargement.py

Charge le fichier CSV dans un DataFrame pandas.

### nettoyage.py

Nettoie les noms de colonnes et convertit les colonnes numériques.

### analyse.py

Réalise les calculs et produit les indicateurs utilisés dans le tableau de bord.

### export_excel.py

Construit l'ensemble du fichier Excel :

* feuilles de calcul ;
* filtres ;
* KPI ;
* graphiques ;
* tableau de bord.

### main.py

Lance l'ensemble du projet :

* chargement ;
* nettoyage ;
* analyse ;
* génération du fichier Excel.

---

## Bibliothèques utilisées

* pandas
* openpyxl

