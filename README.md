# Excel SQL Manager 🐍📊

Un projet Python puissant pour lire, manipuler et mettre à jour des fichiers Excel en utilisant des requêtes SQL.

## 🎯 Fonctionnalités

- **Lecture de fichiers Excel** : Chargement automatique de toutes les feuilles
- **Requêtes SQL** : Exécution de requêtes SQL directement sur les données Excel
- **Manipulation de données** : Filtrage, groupement, jointures et analyses avancées
- **Mise à jour Excel** : Sauvegarde des résultats dans des fichiers Excel
- **Sauvegarde automatique** : Protection des données originales
- **Analyses prédéfinies** : Requêtes courantes pour l'analyse de données
- **Rapport de qualité** : Évaluation automatique de la qualité des données

## 🚀 Installation

1. Clonez ce repository :
```bash
git clone https://github.com/Bernard-Vesy/Python-ExcelFile-SQL.git
cd Python-ExcelFile-SQL
```

2. Créez un environnement virtuel (recommandé) :
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# ou
source .venv/bin/activate  # Linux/Mac
```

3. Installez les dépendances :
```bash
pip install -r requirements.txt
```

## 📁 Structure du projet

```
Python-ExcelFile-SQL/
├── excel_sql_manager.py    # Classe principale ExcelSQLManager
├── sql_utils.py           # Utilitaires et requêtes SQL prédéfinies
├── example_usage.py       # Exemples d'utilisation
├── requirements.txt       # Dépendances Python
├── README.md             # Documentation
└── fiches_lemo_extended.xlsx  # Fichier Excel d'exemple
```

## 🔧 Utilisation de base

### Exemple simple

```python
from excel_sql_manager import ExcelSQLManager

# Utilisation avec context manager (recommandé)
with ExcelSQLManager("mon_fichier.xlsx") as manager:
    # Charger les feuilles Excel comme tables SQL
    manager.load_sheets_to_sql()
    
    # Exécuter une requête SQL
    result = manager.execute_query("SELECT * FROM Sheet1 WHERE colonne > 100")
    print(result)
    
    # Mettre à jour une feuille avec une requête
    manager.update_sheet_from_query("Sheet1", 
        "SELECT *, colonne * 2 as colonne_double FROM Sheet1")
    
    # Sauvegarder
    manager.save_excel("fichier_modifie.xlsx")
```

### Utilisation des fonctions utilitaires

```python
from excel_sql_manager import quick_excel_query, update_excel_with_query

# Requête rapide
result = quick_excel_query("fichier.xlsx", "SELECT COUNT(*) FROM Sheet1")

# Mise à jour rapide
update_excel_with_query("fichier.xlsx", "Sheet1", 
                       "SELECT * FROM Sheet1 WHERE statut = 'actif'",
                       backup=True)
```

## 🛠️ Fonctionnalités avancées

### Constructeur de requêtes SQL

```python
from sql_utils import SQLQueryBuilder

# Construction automatique de requêtes
query = SQLQueryBuilder.filter_by_value("ma_table", "prix", 100, ">")
query = SQLQueryBuilder.group_by_count("ma_table", "categorie")
query = SQLQueryBuilder.top_n_records("ma_table", "ventes", 10)
```

### Analyses prédéfinies

```python
from sql_utils import DataAnalysisQueries

# Analyse de tendance mensuelle
query = DataAnalysisQueries.monthly_trend("ventes", "date", "montant")

# Analyse de Pareto
query = DataAnalysisQueries.pareto_analysis("produits", "nom", "ventes")

# Détection d'outliers
query = DataAnalysisQueries.outlier_detection("donnees", "valeur")
```

### Rapport de qualité des données

```python
from sql_utils import generate_data_quality_report

with ExcelSQLManager("fichier.xlsx") as manager:
    manager.load_sheets_to_sql()
    rapport = generate_data_quality_report(manager, "Sheet1")
    print(f"Complétude des données: {rapport['columns_stats']}")
```

## 📊 Exemples d'analyses courantes

### 1. Analyse de ventes par mois
```python
with ExcelSQLManager("ventes.xlsx") as manager:
    manager.load_sheets_to_sql()
    
    monthly_sales = manager.execute_query("""
        SELECT 
            strftime('%Y-%m', date) as mois,
            SUM(montant) as total_ventes,
            COUNT(*) as nombre_ventes
        FROM ventes 
        GROUP BY strftime('%Y-%m', date)
        ORDER BY mois
    """)
    print(monthly_sales)
```

### 2. Top 10 des clients
```python
top_clients = manager.execute_query("""
    SELECT 
        client,
        SUM(montant) as total_achats,
        COUNT(*) as nombre_commandes
    FROM commandes 
    GROUP BY client
    ORDER BY total_achats DESC
    LIMIT 10
""")
```

### 3. Détection de doublons
```python
doublons = manager.execute_query("""
    SELECT email, COUNT(*) as occurrences
    FROM clients
    GROUP BY email
    HAVING COUNT(*) > 1
""")
```

## 🧪 Tests et exemples

Exécutez le script d'exemples pour voir toutes les fonctionnalités en action :

```bash
python example_usage.py
```

Ce script :
- Charge un fichier Excel d'exemple
- Montre différents types de requêtes SQL
- Génère un rapport de qualité des données
- Crée des fichiers modifiés
- Démontre les fonctionnalités avancées

## 📋 API Reference

### ExcelSQLManager

#### Méthodes principales :
- `load_excel_to_memory()` : Charge le fichier Excel en mémoire
- `load_sheets_to_sql(sheet_names=None)` : Charge les feuilles comme tables SQL
- `execute_query(query)` : Exécute une requête SELECT
- `execute_update(query)` : Exécute UPDATE/INSERT/DELETE
- `update_sheet_from_query(sheet_name, query)` : Met à jour une feuille
- `save_excel(output_path=None)` : Sauvegarde le fichier Excel
- `backup_original()` : Crée une sauvegarde
- `get_table_info(table_name)` : Informations sur une table
- `list_tables()` : Liste toutes les tables disponibles

### SQLQueryBuilder

#### Méthodes utiles :
- `select_all(table)` : SELECT *
- `filter_by_value(table, column, value, operator)` : Filtrage
- `group_by_count(table, column)` : Groupement avec comptage
- `join_tables(table1, table2, join_column)` : Jointures
- `find_duplicates(table, columns)` : Recherche de doublons
- `basic_statistics(table, column)` : Statistiques de base

## 🔧 Configuration avancée

### Base de données personnalisée

```python
# Utiliser une base SQLite permanente
manager = ExcelSQLManager("fichier.xlsx", db_path="ma_base.db")
```

### Gestion des erreurs

```python
try:
    with ExcelSQLManager("fichier.xlsx") as manager:
        manager.load_sheets_to_sql()
        result = manager.execute_query("SELECT * FROM table_inexistante")
except FileNotFoundError:
    print("Fichier Excel non trouvé")
except Exception as e:
    print(f"Erreur: {e}")
```

## 🐛 Résolution de problèmes

### Problèmes courants :

1. **"Table doesn't exist"** : Vérifiez que `load_sheets_to_sql()` a été appelé
2. **Noms de colonnes avec espaces** : Utilisez des guillemets : `"nom colonne"`
3. **Caractères spéciaux** : Les noms de tables sont automatiquement nettoyés
4. **Fichiers Excel corrompus** : Vérifiez l'intégrité du fichier

### Debugging :

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## 📄 Licence

Ce projet est sous licence MIT. Voir le fichier LICENSE pour plus de détails.

## 🤝 Contribution

Les contributions sont les bienvenues ! 

1. Fork le projet
2. Créez une branche pour votre fonctionnalité
3. Committez vos changements
4. Poussez vers la branche
5. Ouvrez une Pull Request

## 📞 Support

Pour toute question ou problème :
- Ouvrez une issue sur GitHub
- Consultez les exemples dans `example_usage.py`
- Vérifiez la documentation dans le code

## 🚀 Roadmap

- [ ] Support pour d'autres formats (CSV, JSON)
- [ ] Interface graphique
- [ ] Requêtes SQL plus complexes (window functions)
- [ ] Export vers différents formats
- [ ] Intégration avec des APIs
- [ ] Tests unitaires automatisés

---

**Créé avec ❤️ et GitHub Copilot**