# Accès à l'eau potable dans le monde : DWFA

> Dashboard Tableau et analyse géographique de l'accès à l'eau dans 194 pays (2000–2017), pour identifier dans chaque domaine d'intervention le pays le plus en difficulté.

**Stack** · Tableau · Python · pandas · Jupyter

---

## Contexte

DWFA (Drinking Water For All) est une mission d'analyse sur l'accès à l'eau potable dans le monde.
Objectif : cerner, dans chacun de trois domaines d'expertise, le pays le plus en difficulté, pour orienter l'intervention.

| Domaine | Situation visée | Critère |
|---------|-----------------|---------|
| Création de services | Pas d'infrastructure de base | Population non couverte |
| Modernisation des services | Infrastructure partielle | Écart entre accès basique et accès sécurisé |
| Conseil (consulting) | Instabilité politique élevée | Score de stabilité politique |

## Données

5 fichiers sources : accès à l'eau (basique et sécurisé), mortalité liée à l'eau et à l'assainissement (WASH), population, stabilité politique, régions.
Ils sont consolidés en un fichier unique : **3 492 lignes, 13 colonnes, 194 pays, 18 années (2000–2017), 6 régions, 0 doublon**.

Lacunes assumées et documentées dans le notebook :

- l'accès aux services sécurisés n'est renseigné que pour 98 pays sur 194 (50 % de valeurs manquantes)
- la mortalité WASH n'existe que pour une seule année, 2016 (183 pays)
- la part de population urbaine dépasse 100 % pour quelques pays (Koweït, Singapour, Nauru, Monaco) : le notebook l'attribue à des travailleurs migrants comptés au-delà de la population résidente dans la source

## Méthodologie

1. **Blueprint et mock-up** : besoins utilisateurs traduits en mesures, visualisations et pages (`Etude_sur_l_eau_potable-Blueprint.pdf`)
2. **Nettoyage et fusion** des 5 fichiers (`pretraitement_DWFA.ipynb`) : renommage, typage, jointures, tests de validation automatisés
3. **Enrichissement** : taux de population urbaine et rurale, nombre de décès WASH estimé (taux de mortalité × population)
4. **Vues globales** dans Tableau : mondiale, continentale, nationale
5. **Vues par domaine** : trois nuages de points, un par domaine d'intervention

## Résultats clés

Chiffres recalculés à partir de `dwfa_consolide.csv` :

- La population sans accès basique à l'eau passe d'environ **1,11 milliard de personnes en 2000 à environ 751 millions en 2017** (calcul sur les 177 pays renseignés les deux années)
- En 2017, **l'Afrique concentre environ la moitié** de la population sans accès basique (environ 387 millions sur 778 millions)
- En 2016, l'Afrique compte environ **464 000 des 868 000 décès WASH estimés** (53 %)
- Synthèse de la présentation : les pays d'Afrique sont les plus en difficulté ; les domaines création et modernisation concernent surtout les pays en développement

## Limites

- Le décompte des décès WASH est une estimation (taux × population), sur une seule année
- Les comparaisons dans le temps portent sur les pays qui ont des données les deux années
- L'accès sécurisé n'est renseigné que pour 98 pays, ce qui limite l'analyse du domaine modernisation
- Le domaine conseil repose sur le seul score de stabilité politique

## Contenu du repo

| Fichier | Description |
|---------|-------------|
| `pretraitement_DWFA.ipynb` | Nettoyage, fusion, indicateurs dérivés, tests de validation |
| `dwfa_consolide.csv` | Table principale : un pays par année (3 492 lignes) |
| `dwfa_eau_granulaire.csv` | Accès à l'eau en zone urbaine et rurale (6 984 lignes) |
| `dashboard_eau_potable_FINAL.twb` | Classeur Tableau final : 3 dashboards, 17 feuilles |
| `Elices_Diez_Josy_2_dashboard_052026.twb` | Version précédente du classeur (16 feuilles) |
| `Etude_sur_l_eau_potable-Blueprint.pdf` | Blueprint : besoins, mesures, visualisations |
| `Etude_sur_l_eau_potable-Présentation.pdf` | Présentation de l'étude (25 diapositives) |
| `Elices_Diez_Josy_1_presentation_052026.pptx` | Support de présentation |

## Reproduire

Les 5 fichiers sources bruts ne sont pas inclus ; les deux CSV produits par le notebook le sont.
Pour explorer le dashboard, ouvrir un `.twb` dans Tableau Desktop ou Tableau Public : le classeur pointe vers un chemin local, Tableau demande de repointer `dwfa_consolide.csv` et `dwfa_eau_granulaire.csv` vers les fichiers du repo.

---

*Formation Data Analyst OpenClassrooms × ENSAE*
