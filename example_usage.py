"""
Exemples d'utilisation du projet Excel-SQL Manager
Author: Created with GitHub Copilot
Date: November 2025

Ce script montre différentes façons d'utiliser le ExcelSQLManager pour 
manipuler des fichiers Excel avec des requêtes SQL.
"""

from excel_sql_manager import ExcelSQLManager, quick_excel_query, update_excel_with_query
from sql_utils import SQLQueryBuilder, DataAnalysisQueries, generate_data_quality_report, suggest_queries_for_table
import pandas as pd
import numpy as np
from pathlib import Path


def exemple_basique():
    """
    Exemple basique : charger un fichier Excel et exécuter des requêtes simples.
    """
    print("=" * 60)
    print("EXEMPLE 1: Utilisation basique")
    print("=" * 60)
    
    excel_file = "fiches_lemo_extended.xlsx"
    
    # Vérifier que le fichier existe
    if not Path(excel_file).exists():
        print(f"⚠️  Le fichier {excel_file} n'existe pas.")
        print("Créons un fichier d'exemple pour la démonstration...")
        creer_fichier_exemple()
        return
    
    # Utiliser le context manager pour une gestion automatique des ressources
    with ExcelSQLManager(excel_file) as manager:
        print(f"📁 Chargement du fichier: {excel_file}")
        
        # Charger toutes les feuilles en mémoire
        excel_data = manager.load_excel_to_memory()
        print(f"📊 Feuilles trouvées: {list(excel_data.keys())}")
        
        # Charger les feuilles comme tables SQL
        manager.load_sheets_to_sql()
        
        # Lister toutes les tables disponibles
        tables = manager.list_tables()
        print(f"🗄️  Tables SQL créées: {tables}")
        
        # Pour chaque table, afficher un aperçu
        for table in tables:
            print(f"\n--- Aperçu de la table '{table}' ---")
            preview = manager.preview_data(table, limit=3)
            print(preview)
            
            # Informations sur la table
            info = manager.get_table_info(table)
            print(f"📈 Nombre de lignes: {info['row_count']}")
            print(f"📋 Colonnes: {[col['name'] for col in info['columns']]}")


def exemple_requetes_sql():
    """
    Exemple d'utilisation de requêtes SQL avec le constructeur de requêtes.
    """
    print("\n" + "=" * 60)
    print("EXEMPLE 2: Requêtes SQL avec SQLQueryBuilder")
    print("=" * 60)
    
    excel_file = "fiches_lemo_extended.xlsx"
    
    if not Path(excel_file).exists():
        print(f"⚠️  Le fichier {excel_file} n'existe pas.")
        return
    
    with ExcelSQLManager(excel_file) as manager:
        manager.load_sheets_to_sql()
        tables = manager.list_tables()
        
        if not tables:
            print("Aucune table trouvée.")
            return
        
        table_name = tables[0]  # Prendre la première table
        print(f"🎯 Analyse de la table: {table_name}")
        
        # Exemple 1: Sélection de toutes les données
        print("\n1. Toutes les données (5 premières lignes):")
        query = SQLQueryBuilder.select_all(table_name) + " LIMIT 5"
        result = manager.execute_query(query)
        print(result)
        
        # Exemple 2: Informations sur la table
        info = manager.get_table_info(table_name)
        columns = [col['name'] for col in info['columns']]
        print(f"\n2. Colonnes disponibles: {columns}")
        
        # Exemple 3: Statistiques de base si on trouve une colonne numérique
        for col in columns:
            # Essayer de détecter des colonnes numériques
            sample_query = f"SELECT {col} FROM {table_name} WHERE {col} IS NOT NULL LIMIT 1"
            try:
                sample = manager.execute_query(sample_query)
                if not sample.empty:
                    sample_value = sample.iloc[0, 0]
                    if isinstance(sample_value, (int, float)):
                        print(f"\n3. Statistiques pour la colonne numérique '{col}':")
                        stats_query = SQLQueryBuilder.basic_statistics(table_name, col)
                        stats = manager.execute_query(stats_query)
                        print(stats)
                        break
            except:
                continue
        
        # Exemple 4: Groupement par une colonne (si possible)
        if len(columns) > 1:
            group_col = columns[0]
            print(f"\n4. Distribution des valeurs pour '{group_col}':")
            group_query = SQLQueryBuilder.group_by_count(table_name, group_col)
            try:
                grouped = manager.execute_query(group_query)
                print(grouped.head())
            except Exception as e:
                print(f"Erreur lors du groupement: {e}")


def exemple_analyse_donnees():
    """
    Exemple d'analyse de données avec des requêtes prédéfinies.
    """
    print("\n" + "=" * 60)
    print("EXEMPLE 3: Analyse de données avancée")
    print("=" * 60)
    
    excel_file = "fiches_lemo_extended.xlsx"
    
    if not Path(excel_file).exists():
        print(f"⚠️  Le fichier {excel_file} n'existe pas.")
        return
    
    with ExcelSQLManager(excel_file) as manager:
        manager.load_sheets_to_sql()
        tables = manager.list_tables()
        
        if not tables:
            print("Aucune table trouvée.")
            return
        
        table_name = tables[0]
        print(f"🔍 Analyse de qualité des données pour: {table_name}")
        
        # Générer un rapport de qualité des données
        try:
            rapport = generate_data_quality_report(manager, table_name)
            
            print(f"\n📊 Rapport de qualité:")
            print(f"   • Nombre total de lignes: {rapport['table_info']['row_count']}")
            print(f"   • Nombre de colonnes: {len(rapport['columns_stats'])}")
            
            print(f"\n📋 Qualité par colonne:")
            for col_name, stats in rapport['columns_stats'].items():
                print(f"   • {col_name}:")
                print(f"     - Complétude: {stats['completeness']}%")
                print(f"     - Valeurs uniques: {stats['unique_count']}")
                print(f"     - Valeurs nulles: {stats['null_count']}")
                
        except Exception as e:
            print(f"Erreur lors de la génération du rapport: {e}")
        
        # Suggestions de requêtes
        print(f"\n💡 Suggestions de requêtes pour cette table:")
        try:
            suggestions = suggest_queries_for_table(manager, table_name)
            for i, suggestion in enumerate(suggestions[:5], 1):  # Limiter à 5 suggestions
                print(f"\n{i}. {suggestion['description']}")
                print(f"   Requête: {suggestion['query']}")
                
                # Exécuter la première suggestion comme exemple
                if i == 1:
                    try:
                        result = manager.execute_query(suggestion['query'])
                        print(f"   Résultat:")
                        print(result)
                    except Exception as e:
                        print(f"   Erreur: {e}")
                        
        except Exception as e:
            print(f"Erreur lors de la génération des suggestions: {e}")


def exemple_modification_donnees():
    """
    Exemple de modification de données et sauvegarde.
    """
    print("\n" + "=" * 60)
    print("EXEMPLE 4: Modification et sauvegarde de données")
    print("=" * 60)
    
    excel_file = "fiches_lemo_extended.xlsx"
    
    if not Path(excel_file).exists():
        print(f"⚠️  Le fichier {excel_file} n'existe pas.")
        creer_fichier_exemple()
        excel_file = "exemple_donnees.xlsx"
    
    with ExcelSQLManager(excel_file) as manager:
        # Créer une sauvegarde avant modification
        print("💾 Création d'une sauvegarde...")
        try:
            manager.backup_original()
            print("✅ Sauvegarde créée")
        except Exception as e:
            print(f"❌ Erreur lors de la sauvegarde: {e}")
        
        manager.load_sheets_to_sql()
        tables = manager.list_tables()
        
        if not tables:
            print("Aucune table trouvée.")
            return
        
        table_name = tables[0]
        sheet_name = list(manager.original_excel_data.keys())[0]
        
        print(f"🔄 Modification de la feuille: {sheet_name}")
        
        # Exemple: créer une nouvelle vue des données avec une requête
        # Par exemple, ajouter une colonne calculée ou filtrer les données
        modification_query = f"""
        SELECT *,
               'Processé le {pd.Timestamp.now().strftime("%Y-%m-%d")}' as date_traitement
        FROM {table_name}
        """
        
        try:
            # Mettre à jour la feuille avec la requête modifiée
            manager.update_sheet_from_query(sheet_name, modification_query)
            
            # Sauvegarder dans un nouveau fichier
            output_file = f"resultat_modifie_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            manager.save_excel(output_file)
            
            print(f"✅ Données modifiées et sauvegardées dans: {output_file}")
            
        except Exception as e:
            print(f"❌ Erreur lors de la modification: {e}")


def exemple_fonctions_utilitaires():
    """
    Exemple d'utilisation des fonctions utilitaires rapides.
    """
    print("\n" + "=" * 60)
    print("EXEMPLE 5: Fonctions utilitaires rapides")
    print("=" * 60)
    
    excel_file = "fiches_lemo_extended.xlsx"
    
    if not Path(excel_file).exists():
        print(f"⚠️  Le fichier {excel_file} n'existe pas.")
        return
    
    # Exemple 1: Exécution rapide d'une requête
    print("🚀 Exécution rapide d'une requête:")
    try:
        result = quick_excel_query(excel_file, "SELECT * FROM Sheet1 LIMIT 3")
        print("Résultat de la requête rapide:")
        print(result)
    except Exception as e:
        print(f"Erreur: {e}")
        
        # Essayer avec une table différente
        try:
            with ExcelSQLManager(excel_file) as manager:
                manager.load_sheets_to_sql()
                tables = manager.list_tables()
                if tables:
                    table_name = tables[0]
                    result = quick_excel_query(excel_file, f"SELECT * FROM {table_name} LIMIT 3")
                    print(f"Résultat de la requête rapide sur {table_name}:")
                    print(result)
        except Exception as e2:
            print(f"Erreur secondaire: {e2}")


def creer_fichier_exemple():
    """
    Crée un fichier Excel d'exemple pour les démonstrations.
    """
    print("📝 Création d'un fichier Excel d'exemple...")
    
    # Données d'exemple
    donnees_ventes = {
        'Date': pd.date_range('2024-01-01', periods=50, freq='D'),
        'Produit': (['Produit A', 'Produit B', 'Produit C'] * 16 + ['Produit A', 'Produit B'])[:50],
        'Quantite': np.random.randint(1, 100, 50),
        'Prix_Unitaire': np.random.uniform(10, 500, 50).round(2),
        'Vendeur': (['Alice', 'Bob', 'Charlie', 'Diana'] * 12 + ['Alice', 'Bob'])[:50]
    }
    
    donnees_clients = {
        'ID_Client': range(1, 21),
        'Nom': [f'Client_{i}' for i in range(1, 21)],
        'Ville': ['Paris', 'Lyon', 'Marseille', 'Toulouse', 'Nice'] * 4,
        'Age': np.random.randint(18, 80, 20),
        'Statut': ['Actif', 'Inactif'] * 10
    }
    
    # Calculer le montant total
    df_ventes = pd.DataFrame(donnees_ventes)
    df_ventes['Montant_Total'] = df_ventes['Quantite'] * df_ventes['Prix_Unitaire']
    
    df_clients = pd.DataFrame(donnees_clients)
    
    # Sauvegarder dans Excel
    with pd.ExcelWriter('exemple_donnees.xlsx', engine='openpyxl') as writer:
        df_ventes.to_excel(writer, sheet_name='Ventes', index=False)
        df_clients.to_excel(writer, sheet_name='Clients', index=False)
    
    print("✅ Fichier d'exemple créé: exemple_donnees.xlsx")
    return 'exemple_donnees.xlsx'


def main():
    """
    Fonction principale qui exécute tous les exemples.
    """
    print("🐍 Excel-SQL Manager - Exemples d'utilisation")
    print("=" * 60)
    
    try:
        # Exécuter tous les exemples
        exemple_basique()
        exemple_requetes_sql()
        exemple_analyse_donnees()
        exemple_modification_donnees()
        exemple_fonctions_utilitaires()
        
        print("\n" + "=" * 60)
        print("✅ Tous les exemples ont été exécutés!")
        print("=" * 60)
        
    except KeyboardInterrupt:
        print("\n❌ Exécution interrompue par l'utilisateur.")
    except Exception as e:
        print(f"\n❌ Erreur inattendue: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()