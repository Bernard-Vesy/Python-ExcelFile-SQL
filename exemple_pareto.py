"""
Exemple d'analyse de Pareto avec Excel-SQL Manager
Démonstration de la règle 80/20 sur des données réelles
"""

from excel_sql_manager import ExcelSQLManager
from sql_utils import DataAnalysisQueries
import pandas as pd

def exemple_pareto_ventes():
    """
    Exemple d'analyse de Pareto sur les données de ventes.
    Identifie quels produits génèrent 80% du chiffre d'affaires.
    """
    print("🎯 ANALYSE DE PARETO - Données de Ventes")
    print("=" * 60)
    
    # Utiliser le fichier d'exemple avec des données de ventes
    with ExcelSQLManager('exemple_donnees.xlsx') as manager:
        manager.load_sheets_to_sql()
        
        # 1. Analyse de Pareto par Produit (CA)
        print("📊 Analyse de Pareto : Quels produits génèrent 80% du CA ?")
        pareto_query = DataAnalysisQueries.pareto_analysis(
            'Ventes', 'Produit', 'Montant_Total'
        )
        
        pareto_result = manager.execute_query(pareto_query)
        print(pareto_result)
        
        # 2. Analyse de Pareto par Vendeur
        print("\n👤 Analyse de Pareto : Quels vendeurs génèrent 80% du CA ?")
        pareto_vendeur_query = DataAnalysisQueries.pareto_analysis(
            'Ventes', 'Vendeur', 'Montant_Total'
        )
        
        pareto_vendeur_result = manager.execute_query(pareto_vendeur_query)
        print(pareto_vendeur_result)
        
        # 3. Résumé des insights
        print("\n💡 INSIGHTS DE L'ANALYSE DE PARETO:")
        
        # Produits Top 80%
        top_produits = pareto_result[pareto_result['pareto_category'] == 'Top 80%']
        print(f"   • {len(top_produits)} produits génèrent 80% du CA")
        print(f"   • Produits les plus rentables: {', '.join(top_produits['Produit'].tolist())}")
        
        # Vendeurs Top 80%
        top_vendeurs = pareto_vendeur_result[pareto_vendeur_result['pareto_category'] == 'Top 80%']
        print(f"   • {len(top_vendeurs)} vendeurs génèrent 80% du CA")
        print(f"   • Vendeurs top performers: {', '.join(top_vendeurs['Vendeur'].tolist())}")


def exemple_pareto_composants():
    """
    Exemple d'analyse de Pareto sur les données techniques (composants).
    Identifie quels types de composants sont les plus courants.
    """
    print("\n" + "=" * 60)
    print("🔧 ANALYSE DE PARETO - Données Techniques")
    print("=" * 60)
    
    with ExcelSQLManager('fiches_lemo_extended.xlsx') as manager:
        manager.load_sheets_to_sql()
        
        # 1. Analyse de Pareto par Gender (type de composant)
        print("⚙️  Analyse de Pareto : Quels types de Gender sont les plus courants ?")
        
        # D'abord créer une requête pour compter les occurrences
        count_query = """
        SELECT 
            Gender,
            COUNT(*) as count_total,
            SUM(COUNT(*)) OVER () as grand_total
        FROM Sheet
        GROUP BY Gender
        """
        
        # Puis appliquer l'analyse de Pareto sur ces comptes
        pareto_gender_query = """
        WITH ranked_data AS (
            SELECT 
                Gender,
                COUNT(*) as total_value,
                SUM(COUNT(*)) OVER () as grand_total
            FROM Sheet
            GROUP BY Gender
        ),
        cumulative_data AS (
            SELECT 
                Gender,
                total_value,
                grand_total,
                SUM(total_value) OVER (ORDER BY total_value DESC) as cumulative_value,
                ROUND(100.0 * SUM(total_value) OVER (ORDER BY total_value DESC) / grand_total, 2) as cumulative_percentage
            FROM ranked_data
        )
        SELECT 
            Gender,
            total_value,
            cumulative_value,
            cumulative_percentage,
            CASE WHEN cumulative_percentage <= 80 THEN 'Top 80%' ELSE 'Bottom 20%' END as pareto_category
        FROM cumulative_data
        ORDER BY total_value DESC
        """
        
        pareto_gender_result = manager.execute_query(pareto_gender_query)
        print(pareto_gender_result)
        
        # 2. Analyse de Pareto par Weight (poids des composants)
        print("\n⚖️  Analyse de Pareto : Quels composants contribuent le plus au poids total ?")
        pareto_weight_query = DataAnalysisQueries.pareto_analysis(
            'Sheet', 'Gender', 'Weight'
        )
        
        pareto_weight_result = manager.execute_query(pareto_weight_query)
        print(pareto_weight_result)
        
        # 3. Insights
        print("\n💡 INSIGHTS DE L'ANALYSE TECHNIQUE:")
        
        # Types les plus courants
        top_gender_count = pareto_gender_result[pareto_gender_result['pareto_category'] == 'Top 80%']
        print(f"   • {len(top_gender_count)} types de Gender représentent 80% des composants")
        
        # Types les plus lourds
        top_gender_weight = pareto_weight_result[pareto_weight_result['pareto_category'] == 'Top 80%']
        print(f"   • {len(top_gender_weight)} types de Gender représentent 80% du poids total")


def exemple_pareto_personnalise():
    """
    Exemple d'analyse de Pareto personnalisée.
    Montre comment créer ses propres analyses.
    """
    print("\n" + "=" * 60)
    print("🎨 ANALYSE DE PARETO - Personnalisée")
    print("=" * 60)
    
    with ExcelSQLManager('exemple_donnees.xlsx') as manager:
        manager.load_sheets_to_sql()
        
        # Analyse personnalisée : Pareto par Ville des clients vs Age moyen
        print("🏙️  Analyse personnalisée : Pareto des villes par âge moyen")
        
        pareto_custom_query = """
        WITH ranked_data AS (
            SELECT 
                Ville,
                AVG(CAST(Age AS FLOAT)) as total_value,
                SUM(AVG(CAST(Age AS FLOAT))) OVER () as grand_total
            FROM Clients
            GROUP BY Ville
        ),
        cumulative_data AS (
            SELECT 
                Ville,
                total_value,
                grand_total,
                SUM(total_value) OVER (ORDER BY total_value DESC) as cumulative_value,
                ROUND(100.0 * SUM(total_value) OVER (ORDER BY total_value DESC) / grand_total, 2) as cumulative_percentage
            FROM ranked_data
        )
        SELECT 
            Ville,
            ROUND(total_value, 2) as age_moyen,
            ROUND(cumulative_value, 2) as age_cumule,
            cumulative_percentage,
            CASE WHEN cumulative_percentage <= 80 THEN 'Top 80%' ELSE 'Bottom 20%' END as pareto_category
        FROM cumulative_data
        ORDER BY total_value DESC
        """
        
        pareto_custom_result = manager.execute_query(pareto_custom_query)
        print(pareto_custom_result)
        
        print("\n💡 Cette analyse montre les villes avec les clients les plus âgés")


def guide_utilisation_pareto():
    """
    Guide d'utilisation de l'analyse de Pareto.
    """
    print("\n" + "=" * 60)
    print("📚 GUIDE D'UTILISATION - Analyse de Pareto")
    print("=" * 60)
    
    print("""
🎯 PRINCIPE DE L'ANALYSE DE PARETO:
   La règle 80/20 : souvent 80% des effets proviennent de 20% des causes.
   
📊 UTILISATION AVEC DataAnalysisQueries.pareto_analysis():
   
   pareto_query = DataAnalysisQueries.pareto_analysis(
       table_name='ma_table',
       category_column='categorie',  # Colonne à analyser
       value_column='valeur'         # Colonne des valeurs à sommer
   )
   
🔍 COLONNES DU RÉSULTAT:
   • category_column : La catégorie analysée
   • total_value : Valeur totale pour cette catégorie
   • cumulative_value : Valeur cumulative
   • cumulative_percentage : Pourcentage cumulatif
   • pareto_category : 'Top 80%' ou 'Bottom 20%'
   
💡 EXEMPLES D'APPLICATIONS:
   ✅ Ventes : Quels produits génèrent 80% du CA ?
   ✅ Clients : Quels clients représentent 80% des commandes ?
   ✅ Erreurs : Quels types d'erreurs causent 80% des problèmes ?
   ✅ Stocks : Quels articles représentent 80% de la valeur ?
   
⚠️  CONSEILS:
   • La colonne 'value' doit être numérique
   • Plus il y a de données, plus l'analyse est précise
   • Concentrez-vous sur les éléments 'Top 80%'
    """)


def main():
    """
    Fonction principale qui exécute tous les exemples de Pareto.
    """
    print("📈 EXEMPLES D'ANALYSE DE PARETO - Excel SQL Manager")
    print("=" * 70)
    
    try:
        # Vérifier que les fichiers existent
        from pathlib import Path
        if not Path('exemple_donnees.xlsx').exists():
            print("⚠️  Création du fichier d'exemple...")
            from example_usage import creer_fichier_exemple
            creer_fichier_exemple()
        
        # Exécuter les exemples
        exemple_pareto_ventes()
        exemple_pareto_composants()
        exemple_pareto_personnalise()
        guide_utilisation_pareto()
        
        print("\n" + "=" * 70)
        print("✅ Tous les exemples d'analyse de Pareto terminés !")
        print("=" * 70)
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()