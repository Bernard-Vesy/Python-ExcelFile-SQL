"""
Script de test simple pour le fichier Excel existant
"""

from excel_sql_manager import ExcelSQLManager
from sql_utils import SQLQueryBuilder

def test_simple():
    """Test simple avec le fichier Excel existant"""
    print("🧪 Test du projet Excel-SQL Manager")
    print("=" * 50)
    
    with ExcelSQLManager("fiches_lemo_extended.xlsx") as manager:
        # Charger les données
        manager.load_sheets_to_sql()
        tables = manager.list_tables()
        print(f"📊 Tables trouvées: {tables}")
        
        # Test de quelques requêtes simples
        table_name = tables[0] if tables else None
        if table_name:
            # 1. Nombre total de lignes
            count_query = f"SELECT COUNT(*) as total FROM {table_name}"
            result = manager.execute_query(count_query)
            print(f"📈 Nombre total de lignes: {result['total'].iloc[0]}")
            
            # 2. Distribution par Gender
            gender_query = SQLQueryBuilder.group_by_count(table_name, "Gender")
            gender_result = manager.execute_query(gender_query)
            print(f"\n👫 Distribution par Gender:")
            print(gender_result)
            
            # 3. Statistiques sur NumberContacts
            contacts_query = SQLQueryBuilder.basic_statistics(table_name, "NumberContacts")
            contacts_result = manager.execute_query(contacts_query)
            print(f"\n🔌 Statistiques sur NumberContacts:")
            print(contacts_result)
            
            # 4. Top 5 des produits les plus lourds
            weight_query = SQLQueryBuilder.top_n_records(table_name, "Weight", 5)
            weight_result = manager.execute_query(weight_query)
            print(f"\n⚖️ Top 5 des produits les plus lourds:")
            print(weight_result[['Product', 'Weight', 'Gender', 'NumberContacts']])

if __name__ == "__main__":
    test_simple()