from excel_sql_manager import ExcelSQLManager

print("Test du fichier exemple_donnees.xlsx")
with ExcelSQLManager('exemple_donnees.xlsx') as mgr:
    mgr.load_sheets_to_sql()
    tables = mgr.list_tables()
    print(f'Tables trouvées: {tables}')
    
    for table in tables:
        info = mgr.get_table_info(table)
        print(f'{table}: {info["row_count"]} lignes')
        
        # Aperçu des données
        preview = mgr.preview_data(table, 3)
        print(f"Aperçu de {table}:")
        print(preview)
        print("-" * 50)