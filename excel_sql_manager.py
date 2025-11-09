"""
Excel SQL Manager - Gestionnaire pour lire, manipuler et mettre à jour des fichiers Excel avec SQL
Author: Created with GitHub Copilot
Date: November 2025
"""

import pandas as pd
import sqlite3
from pathlib import Path
from typing import Dict, Any, List, Optional
import logging
from sqlalchemy import create_engine, text
import tempfile
import os

# Configuration du logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ExcelSQLManager:
    """
    Classe principale pour gérer les fichiers Excel avec des opérations SQL.
    
    Cette classe permet de :
    - Lire des fichiers Excel
    - Convertir les feuilles en tables SQL temporaires
    - Exécuter des requêtes SQL sur les données
    - Mettre à jour le fichier Excel avec les résultats
    """
    
    def __init__(self, excel_file_path: str, db_path: Optional[str] = None):
        """
        Initialise le gestionnaire Excel-SQL.
        
        Args:
            excel_file_path (str): Chemin vers le fichier Excel
            db_path (str, optional): Chemin vers la base SQLite. Si None, utilise une DB temporaire
        """
        self.excel_file_path = Path(excel_file_path)
        self.original_excel_data = {}
        self.modified_data = {}
        
        # Configuration de la base de données
        if db_path is None:
            # Créer une base de données temporaire
            self.temp_db = True
            self.db_path = tempfile.mktemp(suffix='.db')
        else:
            self.temp_db = False
            self.db_path = db_path
            
        # Créer l'engine SQLAlchemy
        self.engine = create_engine(f'sqlite:///{self.db_path}')
        self.connection = None
        
        logger.info(f"ExcelSQLManager initialisé avec le fichier: {self.excel_file_path}")
        logger.info(f"Base de données: {self.db_path}")
        
    def __enter__(self):
        """Context manager entry"""
        self.connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()
        
    def connect(self):
        """Établit la connexion à la base de données"""
        try:
            self.connection = self.engine.connect()
            logger.info("Connexion à la base de données établie")
        except Exception as e:
            logger.error(f"Erreur lors de la connexion à la base de données: {e}")
            raise
            
    def close(self):
        """Ferme la connexion et nettoie les ressources"""
        if self.connection:
            self.connection.close()
            logger.info("Connexion fermée")
            
        # Supprimer la base temporaire si nécessaire
        if self.temp_db and os.path.exists(self.db_path):
            try:
                os.remove(self.db_path)
                logger.info("Base de données temporaire supprimée")
            except Exception as e:
                logger.warning(f"Impossible de supprimer la base temporaire: {e}")
    
    def load_excel_to_memory(self):
        """
        Charge toutes les feuilles du fichier Excel en mémoire.
        
        Returns:
            Dict[str, pd.DataFrame]: Dictionnaire des DataFrames par nom de feuille
        """
        try:
            if not self.excel_file_path.exists():
                raise FileNotFoundError(f"Le fichier Excel n'existe pas: {self.excel_file_path}")
                
            # Lire toutes les feuilles
            excel_data = pd.read_excel(self.excel_file_path, sheet_name=None, engine='openpyxl')
            self.original_excel_data = excel_data.copy()
            self.modified_data = excel_data.copy()
            
            logger.info(f"Fichier Excel chargé. Feuilles trouvées: {list(excel_data.keys())}")
            return excel_data
            
        except Exception as e:
            logger.error(f"Erreur lors du chargement du fichier Excel: {e}")
            raise
    
    def load_sheets_to_sql(self, sheet_names: Optional[List[str]] = None):
        """
        Charge les feuilles Excel comme tables SQL dans la base de données.
        
        Args:
            sheet_names (List[str], optional): Liste des noms de feuilles à charger.
                                             Si None, charge toutes les feuilles.
        """
        if not self.original_excel_data:
            self.load_excel_to_memory()
            
        sheets_to_load = sheet_names if sheet_names else list(self.original_excel_data.keys())
        
        for sheet_name in sheets_to_load:
            if sheet_name not in self.original_excel_data:
                logger.warning(f"Feuille '{sheet_name}' non trouvée dans le fichier Excel")
                continue
                
            try:
                df = self.original_excel_data[sheet_name]
                
                # Nettoyer le nom de la table (enlever caractères spéciaux)
                table_name = self._clean_table_name(sheet_name)
                
                # Charger dans SQLite
                df.to_sql(table_name, self.engine, if_exists='replace', index=False)
                logger.info(f"Feuille '{sheet_name}' chargée comme table '{table_name}' ({len(df)} lignes)")
                
            except Exception as e:
                logger.error(f"Erreur lors du chargement de la feuille '{sheet_name}': {e}")
                raise
    
    def execute_query(self, query: str) -> pd.DataFrame:
        """
        Exécute une requête SQL et retourne le résultat sous forme de DataFrame.
        
        Args:
            query (str): Requête SQL à exécuter
            
        Returns:
            pd.DataFrame: Résultat de la requête
        """
        try:
            result = pd.read_sql_query(query, self.engine)
            logger.info(f"Requête exécutée avec succès. Résultat: {len(result)} lignes")
            return result
            
        except Exception as e:
            logger.error(f"Erreur lors de l'exécution de la requête: {e}")
            logger.error(f"Requête: {query}")
            raise
    
    def execute_update(self, query: str) -> int:
        """
        Exécute une requête de mise à jour (INSERT, UPDATE, DELETE).
        
        Args:
            query (str): Requête SQL de mise à jour
            
        Returns:
            int: Nombre de lignes affectées
        """
        try:
            with self.engine.begin() as conn:
                result = conn.execute(text(query))
                affected_rows = result.rowcount
                logger.info(f"Requête de mise à jour exécutée. Lignes affectées: {affected_rows}")
                return affected_rows
                
        except Exception as e:
            logger.error(f"Erreur lors de l'exécution de la mise à jour: {e}")
            logger.error(f"Requête: {query}")
            raise
    
    def get_table_info(self, table_name: str) -> Dict[str, Any]:
        """
        Obtient des informations sur une table SQL.
        
        Args:
            table_name (str): Nom de la table
            
        Returns:
            Dict[str, Any]: Informations sur la table
        """
        try:
            # Informations sur les colonnes
            columns_query = f"PRAGMA table_info({table_name})"
            columns_info = pd.read_sql_query(columns_query, self.engine)
            
            # Nombre de lignes
            count_query = f"SELECT COUNT(*) as row_count FROM {table_name}"
            row_count = pd.read_sql_query(count_query, self.engine)['row_count'].iloc[0]
            
            return {
                'table_name': table_name,
                'columns': columns_info.to_dict('records'),
                'row_count': row_count
            }
            
        except Exception as e:
            logger.error(f"Erreur lors de la récupération des informations de table '{table_name}': {e}")
            raise
    
    def update_sheet_from_query(self, sheet_name: str, query: str):
        """
        Met à jour une feuille Excel avec le résultat d'une requête SQL.
        
        Args:
            sheet_name (str): Nom de la feuille à mettre à jour
            query (str): Requête SQL dont le résultat remplacera la feuille
        """
        try:
            # Exécuter la requête
            new_data = self.execute_query(query)
            
            # Mettre à jour les données en mémoire
            self.modified_data[sheet_name] = new_data
            
            logger.info(f"Feuille '{sheet_name}' mise à jour avec {len(new_data)} lignes")
            
        except Exception as e:
            logger.error(f"Erreur lors de la mise à jour de la feuille '{sheet_name}': {e}")
            raise
    
    def save_excel(self, output_path: Optional[str] = None):
        """
        Sauvegarde les données modifiées dans un fichier Excel.
        
        Args:
            output_path (str, optional): Chemin de sortie. Si None, écrase le fichier original.
        """
        try:
            save_path = output_path if output_path else self.excel_file_path
            
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for sheet_name, df in self.modified_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            logger.info(f"Fichier Excel sauvegardé: {save_path}")
            
        except Exception as e:
            logger.error(f"Erreur lors de la sauvegarde du fichier Excel: {e}")
            raise
    
    def backup_original(self, backup_path: Optional[str] = None):
        """
        Crée une sauvegarde du fichier Excel original.
        
        Args:
            backup_path (str, optional): Chemin de la sauvegarde. 
                                       Si None, ajoute '_backup' au nom du fichier.
        """
        try:
            if backup_path is None:
                backup_path = str(self.excel_file_path.with_name(
                    f"{self.excel_file_path.stem}_backup{self.excel_file_path.suffix}"
                ))
            
            import shutil
            # Convertir en string pour éviter les problèmes avec Path objects
            shutil.copy2(str(self.excel_file_path), backup_path)
            logger.info(f"Sauvegarde créée: {backup_path}")
            
        except Exception as e:
            logger.error(f"Erreur lors de la création de la sauvegarde: {e}")
            raise
    
    def list_tables(self) -> List[str]:
        """
        Liste toutes les tables disponibles dans la base de données.
        
        Returns:
            List[str]: Liste des noms de tables
        """
        try:
            query = "SELECT name FROM sqlite_master WHERE type='table'"
            tables = pd.read_sql_query(query, self.engine)
            return tables['name'].tolist()
            
        except Exception as e:
            logger.error(f"Erreur lors de la récupération de la liste des tables: {e}")
            raise
    
    def preview_data(self, table_name: str, limit: int = 5) -> pd.DataFrame:
        """
        Aperçu des données d'une table.
        
        Args:
            table_name (str): Nom de la table
            limit (int): Nombre de lignes à afficher
            
        Returns:
            pd.DataFrame: Aperçu des données
        """
        query = f"SELECT * FROM {table_name} LIMIT {limit}"
        return self.execute_query(query)
    
    @staticmethod
    def _clean_table_name(name: str) -> str:
        """
        Nettoie un nom pour l'utiliser comme nom de table SQL.
        
        Args:
            name (str): Nom original
            
        Returns:
            str: Nom nettoyé
        """
        # Remplacer les caractères spéciaux par des underscores
        import re
        cleaned = re.sub(r'[^\w]', '_', name)
        # S'assurer que le nom commence par une lettre ou underscore
        if cleaned and not cleaned[0].isalpha() and cleaned[0] != '_':
            cleaned = '_' + cleaned
        return cleaned or 'unnamed_table'


# Fonctions utilitaires
def quick_excel_query(excel_path: str, query: str, sheet_names: Optional[List[str]] = None) -> pd.DataFrame:
    """
    Fonction utilitaire pour exécuter rapidement une requête sur un fichier Excel.
    
    Args:
        excel_path (str): Chemin vers le fichier Excel
        query (str): Requête SQL à exécuter
        sheet_names (List[str], optional): Feuilles à charger
        
    Returns:
        pd.DataFrame: Résultat de la requête
    """
    with ExcelSQLManager(excel_path) as manager:
        manager.load_sheets_to_sql(sheet_names)
        return manager.execute_query(query)


def update_excel_with_query(excel_path: str, sheet_name: str, query: str, 
                          output_path: Optional[str] = None, backup: bool = True):
    """
    Fonction utilitaire pour mettre à jour un fichier Excel avec une requête.
    
    Args:
        excel_path (str): Chemin vers le fichier Excel
        sheet_name (str): Nom de la feuille à mettre à jour
        query (str): Requête SQL pour les nouvelles données
        output_path (str, optional): Chemin de sortie
        backup (bool): Créer une sauvegarde avant modification
    """
    with ExcelSQLManager(excel_path) as manager:
        if backup:
            manager.backup_original()
            
        manager.load_sheets_to_sql()
        manager.update_sheet_from_query(sheet_name, query)
        manager.save_excel(output_path)