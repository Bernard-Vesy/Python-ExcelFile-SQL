"""
SQL Utilities - Fonctions et requêtes SQL courantes pour l'analyse de données Excel
Author: Created with GitHub Copilot
Date: November 2025
"""

from typing import List, Dict, Any, Optional
import pandas as pd
from datetime import datetime, timedelta


class SQLQueryBuilder:
    """
    Constructeur de requêtes SQL pour des opérations courantes sur des données Excel.
    """
    
    @staticmethod
    def select_all(table_name: str) -> str:
        """Sélectionne toutes les données d'une table."""
        return f"SELECT * FROM {table_name}"
    
    @staticmethod
    def select_columns(table_name: str, columns: List[str]) -> str:
        """Sélectionne des colonnes spécifiques."""
        columns_str = ", ".join(columns)
        return f"SELECT {columns_str} FROM {table_name}"
    
    @staticmethod
    def filter_by_value(table_name: str, column: str, value: Any, 
                       operator: str = "=") -> str:
        """Filtre par valeur dans une colonne."""
        if isinstance(value, str):
            value = f"'{value}'"
        return f"SELECT * FROM {table_name} WHERE {column} {operator} {value}"
    
    @staticmethod
    def filter_by_multiple_values(table_name: str, column: str, 
                                values: List[Any]) -> str:
        """Filtre par plusieurs valeurs (IN clause)."""
        if all(isinstance(v, str) for v in values):
            values_str = ", ".join([f"'{v}'" for v in values])
        else:
            values_str = ", ".join([str(v) for v in values])
        return f"SELECT * FROM {table_name} WHERE {column} IN ({values_str})"
    
    @staticmethod
    def filter_by_range(table_name: str, column: str, min_value: Any, 
                       max_value: Any) -> str:
        """Filtre par plage de valeurs."""
        return f"SELECT * FROM {table_name} WHERE {column} BETWEEN {min_value} AND {max_value}"
    
    @staticmethod
    def filter_by_date_range(table_name: str, date_column: str, 
                            start_date: str, end_date: str) -> str:
        """Filtre par plage de dates."""
        return f"""
        SELECT * FROM {table_name} 
        WHERE {date_column} BETWEEN '{start_date}' AND '{end_date}'
        """
    
    @staticmethod
    def group_by_count(table_name: str, group_column: str, 
                      order_by_count: bool = True) -> str:
        """Groupe par colonne avec comptage."""
        query = f"""
        SELECT {group_column}, COUNT(*) as count
        FROM {table_name}
        GROUP BY {group_column}
        """
        if order_by_count:
            query += " ORDER BY count DESC"
        return query
    
    @staticmethod
    def group_by_sum(table_name: str, group_column: str, sum_column: str) -> str:
        """Groupe par colonne avec somme."""
        return f"""
        SELECT {group_column}, SUM({sum_column}) as total
        FROM {table_name}
        GROUP BY {group_column}
        ORDER BY total DESC
        """
    
    @staticmethod
    def group_by_avg(table_name: str, group_column: str, avg_column: str) -> str:
        """Groupe par colonne avec moyenne."""
        return f"""
        SELECT {group_column}, AVG({avg_column}) as average
        FROM {table_name}
        GROUP BY {group_column}
        ORDER BY average DESC
        """
    
    @staticmethod
    def top_n_records(table_name: str, order_column: str, n: int = 10, 
                     ascending: bool = False) -> str:
        """Sélectionne les N premiers enregistrements."""
        order_dir = "ASC" if ascending else "DESC"
        return f"""
        SELECT * FROM {table_name}
        ORDER BY {order_column} {order_dir}
        LIMIT {n}
        """
    
    @staticmethod
    def find_duplicates(table_name: str, columns: List[str]) -> str:
        """Trouve les doublons basés sur des colonnes."""
        columns_str = ", ".join(columns)
        return f"""
        SELECT {columns_str}, COUNT(*) as duplicate_count
        FROM {table_name}
        GROUP BY {columns_str}
        HAVING COUNT(*) > 1
        ORDER BY duplicate_count DESC
        """
    
    @staticmethod
    def find_missing_values(table_name: str, column: str) -> str:
        """Trouve les valeurs manquantes dans une colonne."""
        return f"""
        SELECT * FROM {table_name}
        WHERE {column} IS NULL OR {column} = ''
        """
    
    @staticmethod
    def basic_statistics(table_name: str, numeric_column: str) -> str:
        """Calcule des statistiques de base pour une colonne numérique."""
        return f"""
        SELECT 
            COUNT({numeric_column}) as count,
            MIN({numeric_column}) as minimum,
            MAX({numeric_column}) as maximum,
            AVG({numeric_column}) as moyenne,
            SUM({numeric_column}) as total
        FROM {table_name}
        WHERE {numeric_column} IS NOT NULL
        """
    
    @staticmethod
    def join_tables(table1: str, table2: str, join_column: str, 
                   join_type: str = "INNER") -> str:
        """Joint deux tables."""
        return f"""
        SELECT t1.*, t2.*
        FROM {table1} t1
        {join_type} JOIN {table2} t2 ON t1.{join_column} = t2.{join_column}
        """
    
    @staticmethod
    def pivot_summary(table_name: str, row_column: str, col_column: str, 
                     value_column: str, aggregation: str = "SUM") -> str:
        """Crée un résumé pivot (limité en SQL standard)."""
        return f"""
        SELECT {row_column},
               {aggregation}(CASE WHEN {col_column} = 'value1' THEN {value_column} END) as value1,
               {aggregation}(CASE WHEN {col_column} = 'value2' THEN {value_column} END) as value2
        FROM {table_name}
        GROUP BY {row_column}
        """


class DataAnalysisQueries:
    """
    Requêtes prédéfinies pour l'analyse de données courantes.
    """
    
    @staticmethod
    def monthly_trend(table_name: str, date_column: str, value_column: str) -> str:
        """Analyse de tendance mensuelle."""
        return f"""
        SELECT 
            strftime('%Y-%m', {date_column}) as month,
            COUNT(*) as count,
            SUM({value_column}) as total,
            AVG({value_column}) as average
        FROM {table_name}
        WHERE {date_column} IS NOT NULL
        GROUP BY strftime('%Y-%m', {date_column})
        ORDER BY month
        """
    
    @staticmethod
    def quarterly_analysis(table_name: str, date_column: str, value_column: str) -> str:
        """Analyse trimestrielle."""
        return f"""
        SELECT 
            strftime('%Y', {date_column}) as year,
            CASE 
                WHEN CAST(strftime('%m', {date_column}) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
                WHEN CAST(strftime('%m', {date_column}) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
                WHEN CAST(strftime('%m', {date_column}) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
                ELSE 'Q4'
            END as quarter,
            SUM({value_column}) as total
        FROM {table_name}
        WHERE {date_column} IS NOT NULL
        GROUP BY year, quarter
        ORDER BY year, quarter
        """
    
    @staticmethod
    def pareto_analysis(table_name: str, category_column: str, value_column: str) -> str:
        """Analyse de Pareto (80/20)."""
        return f"""
        WITH ranked_data AS (
            SELECT 
                {category_column},
                SUM({value_column}) as total_value,
                SUM(SUM({value_column})) OVER () as grand_total
            FROM {table_name}
            GROUP BY {category_column}
        ),
        cumulative_data AS (
            SELECT 
                {category_column},
                total_value,
                grand_total,
                SUM(total_value) OVER (ORDER BY total_value DESC) as cumulative_value,
                ROUND(100.0 * SUM(total_value) OVER (ORDER BY total_value DESC) / grand_total, 2) as cumulative_percentage
            FROM ranked_data
        )
        SELECT 
            {category_column},
            total_value,
            cumulative_value,
            cumulative_percentage,
            CASE WHEN cumulative_percentage <= 80 THEN 'Top 80%' ELSE 'Bottom 20%' END as pareto_category
        FROM cumulative_data
        ORDER BY total_value DESC
        """
    
    @staticmethod
    def outlier_detection(table_name: str, numeric_column: str, 
                         std_deviation_threshold: float = 2.0) -> str:
        """Détection des valeurs aberrantes."""
        return f"""
        WITH stats AS (
            SELECT 
                AVG({numeric_column}) as mean_value,
                AVG({numeric_column} * {numeric_column}) - AVG({numeric_column}) * AVG({numeric_column}) as variance
            FROM {table_name}
            WHERE {numeric_column} IS NOT NULL
        )
        SELECT t.*, 
               s.mean_value,
               SQRT(s.variance) as std_dev,
               ABS(t.{numeric_column} - s.mean_value) / SQRT(s.variance) as z_score,
               CASE 
                   WHEN ABS(t.{numeric_column} - s.mean_value) / SQRT(s.variance) > {std_deviation_threshold} 
                   THEN 'Outlier' 
                   ELSE 'Normal' 
               END as outlier_status
        FROM {table_name} t, stats s
        WHERE t.{numeric_column} IS NOT NULL
        ORDER BY z_score DESC
        """


class DataCleaningQueries:
    """
    Requêtes pour le nettoyage et la transformation de données.
    """
    
    @staticmethod
    def remove_duplicates(table_name: str, unique_columns: List[str]) -> str:
        """Supprime les doublons en gardant le premier occurrence."""
        columns_str = ", ".join(unique_columns)
        return f"""
        DELETE FROM {table_name}
        WHERE rowid NOT IN (
            SELECT MIN(rowid)
            FROM {table_name}
            GROUP BY {columns_str}
        )
        """
    
    @staticmethod
    def update_null_values(table_name: str, column: str, default_value: Any) -> str:
        """Met à jour les valeurs NULL avec une valeur par défaut."""
        if isinstance(default_value, str):
            default_value = f"'{default_value}'"
        return f"""
        UPDATE {table_name}
        SET {column} = {default_value}
        WHERE {column} IS NULL OR {column} = ''
        """
    
    @staticmethod
    def standardize_text(table_name: str, column: str, operation: str = "upper") -> str:
        """Standardise le texte (upper, lower, trim)."""
        operation_map = {
            "upper": f"UPPER(TRIM({column}))",
            "lower": f"LOWER(TRIM({column}))",
            "trim": f"TRIM({column})",
            "title": f"UPPER(SUBSTR(TRIM({column}), 1, 1)) || LOWER(SUBSTR(TRIM({column}), 2))"
        }
        
        if operation not in operation_map:
            raise ValueError(f"Operation '{operation}' not supported. Use: {list(operation_map.keys())}")
        
        return f"""
        UPDATE {table_name}
        SET {column} = {operation_map[operation]}
        WHERE {column} IS NOT NULL
        """
    
    @staticmethod
    def remove_special_characters(table_name: str, column: str) -> str:
        """Supprime les caractères spéciaux d'une colonne texte."""
        return f"""
        UPDATE {table_name}
        SET {column} = REPLACE(
            REPLACE(
                REPLACE(
                    REPLACE(
                        REPLACE({column}, '!', ''),
                        '@', ''),
                    '#', ''),
                '$', ''),
            '%', '')
        WHERE {column} IS NOT NULL
        """


# Fonctions utilitaires pour des analyses rapides
def generate_data_quality_report(manager, table_name: str) -> Dict[str, Any]:
    """
    Génère un rapport de qualité des données pour une table.
    
    Args:
        manager: Instance d'ExcelSQLManager
        table_name (str): Nom de la table à analyser
        
    Returns:
        Dict[str, Any]: Rapport de qualité des données
    """
    report = {}
    
    # Informations de base sur la table
    table_info = manager.get_table_info(table_name)
    report['table_info'] = table_info
    
    # Statistiques par colonne
    columns_stats = {}
    for col_info in table_info['columns']:
        col_name = col_info['name']
        
        # Nombre de valeurs non-nulles
        non_null_query = f"SELECT COUNT({col_name}) as non_null_count FROM {table_name} WHERE {col_name} IS NOT NULL AND {col_name} != ''"
        non_null_count = manager.execute_query(non_null_query)['non_null_count'].iloc[0]
        
        # Valeurs uniques
        unique_query = f"SELECT COUNT(DISTINCT {col_name}) as unique_count FROM {table_name}"
        unique_count = manager.execute_query(unique_query)['unique_count'].iloc[0]
        
        columns_stats[col_name] = {
            'type': col_info['type'],
            'non_null_count': non_null_count,
            'null_count': table_info['row_count'] - non_null_count,
            'unique_count': unique_count,
            'completeness': round(non_null_count / table_info['row_count'] * 100, 2) if table_info['row_count'] > 0 else 0
        }
    
    report['columns_stats'] = columns_stats
    
    return report


def suggest_queries_for_table(manager, table_name: str) -> List[Dict[str, str]]:
    """
    Suggère des requêtes utiles pour une table donnée.
    
    Args:
        manager: Instance d'ExcelSQLManager
        table_name (str): Nom de la table
        
    Returns:
        List[Dict[str, str]]: Liste de suggestions de requêtes
    """
    suggestions = []
    
    # Obtenir les informations de la table
    table_info = manager.get_table_info(table_name)
    columns = [col['name'] for col in table_info['columns']]
    
    # Requêtes de base
    suggestions.append({
        'description': 'Aperçu des données',
        'query': SQLQueryBuilder.select_all(table_name) + ' LIMIT 10'
    })
    
    # Pour chaque colonne, suggérer des analyses
    for col in columns:
        # Analyse de distribution
        suggestions.append({
            'description': f'Distribution des valeurs pour {col}',
            'query': SQLQueryBuilder.group_by_count(table_name, col)
        })
        
        # Vérification des valeurs manquantes
        suggestions.append({
            'description': f'Valeurs manquantes dans {col}',
            'query': SQLQueryBuilder.find_missing_values(table_name, col)
        })
    
    # Recherche de doublons sur toutes les colonnes
    suggestions.append({
        'description': 'Recherche de doublons complets',
        'query': SQLQueryBuilder.find_duplicates(table_name, columns)
    })
    
    return suggestions