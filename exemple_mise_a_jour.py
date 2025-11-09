"""
Exemple de mise à jour de données Excel avec SQL
Démonstration de la modification et sauvegarde de fichiers Excel
"""

from excel_sql_manager import ExcelSQLManager
from sql_utils import SQLQueryBuilder
import pandas as pd
from pathlib import Path

def exemple_mise_a_jour_ventes():
    """
    Exemple : Mise à jour des données de ventes avec des calculs et modifications.
    """
    print("🔄 EXEMPLE DE MISE À JOUR - Données de Ventes")
    print("=" * 60)
    
    # S'assurer que le fichier d'exemple existe
    if not Path('exemple_donnees.xlsx').exists():
        print("⚠️  Création du fichier d'exemple...")
        from example_usage import creer_fichier_exemple
        creer_fichier_exemple()
    
    with ExcelSQLManager('exemple_donnees.xlsx') as manager:
        # Créer une sauvegarde avant modification
        print("💾 Création d'une sauvegarde...")
        manager.backup_original()
        
        # Charger les données
        manager.load_sheets_to_sql()
        
        print("📊 État initial des données:")
        initial_data = manager.execute_query("SELECT * FROM Ventes LIMIT 5")
        print(initial_data[['Produit', 'Quantite', 'Prix_Unitaire', 'Montant_Total']])
        
        # 1. MISE À JOUR : Appliquer une remise de 10% sur les gros montants
        print("\n🔄 MODIFICATION 1: Remise de 10% sur les montants > 3000€")
        
        update_remise_query = """
        SELECT 
            Date,
            Produit,
            Quantite,
            Prix_Unitaire,
            Vendeur,
            CASE 
                WHEN Montant_Total > 3000 THEN ROUND(Montant_Total * 0.9, 2)
                ELSE Montant_Total
            END as Montant_Total,
            CASE 
                WHEN Montant_Total > 3000 THEN 'Remise 10% appliquée'
                ELSE 'Prix normal'
            END as Statut_Prix
        FROM Ventes
        ORDER BY Date
        """
        
        # Mettre à jour la feuille Ventes
        manager.update_sheet_from_query('Ventes', update_remise_query)
        
        # Vérifier les modifications en consultant les données modifiées directement
        modified_data = manager.modified_data['Ventes']
        remises_appliquees = modified_data[modified_data['Statut_Prix'] == 'Remise 10% appliquée']
        
        print(f"✅ {len(remises_appliquees)} lignes modifiées avec remise")
        if not remises_appliquees.empty:
            print(remises_appliquees[['Produit', 'Montant_Total', 'Statut_Prix']].head())
        
        # 2. MISE À JOUR : Ajouter des catégories de performance aux vendeurs
        print("\n🔄 MODIFICATION 2: Ajout de catégories de performance")
        
        # D'abord calculer les performances par vendeur
        perf_query = """
        SELECT 
            Vendeur,
            SUM(Montant_Total) as CA_Total,
            COUNT(*) as Nb_Ventes,
            AVG(Montant_Total) as Panier_Moyen
        FROM Ventes
        GROUP BY Vendeur
        """
        
        perf_data = manager.execute_query(perf_query)
        print("📈 Performance par vendeur:")
        print(perf_data)
        
        # Créer une nouvelle feuille avec les données de performance
        update_clients_query = """
        SELECT 
            c.*,
            CASE 
                WHEN c.Age >= 60 THEN 'Senior'
                WHEN c.Age >= 40 THEN 'Adulte'
                WHEN c.Age >= 25 THEN 'Jeune Adulte'
                ELSE 'Jeune'
            END as Categorie_Age,
            CASE 
                WHEN c.Statut = 'Actif' AND c.Age >= 50 THEN 'Client Premium'
                WHEN c.Statut = 'Actif' THEN 'Client Standard'
                ELSE 'Client Inactif'
            END as Segment_Client
        FROM Clients c
        ORDER BY c.Age DESC
        """
        
        manager.update_sheet_from_query('Clients', update_clients_query)
        
        # Vérifier les nouvelles catégories en utilisant les données modifiées
        clients_modifies = manager.modified_data['Clients']
        categories_data = clients_modifies.groupby(['Categorie_Age', 'Segment_Client']).size().reset_index(name='Nombre')
        print("\n👥 Nouvelles catégories de clients:")
        print(categories_data)
        
        # 3. CRÉER UNE NOUVELLE FEUILLE AVEC UN RÉSUMÉ
        print("\n📋 CRÉATION D'UNE FEUILLE RÉSUMÉ")
        
        resume_query = """
        SELECT 
            'Ventes' as Type_Donnee,
            COUNT(*) as Nombre_Lignes,
            SUM(Montant_Total) as Total_CA,
            AVG(Montant_Total) as Panier_Moyen,
            MIN(Date) as Date_Debut,
            MAX(Date) as Date_Fin
        FROM Ventes
        
        UNION ALL
        
        SELECT 
            'Clients' as Type_Donnee,
            COUNT(*) as Nombre_Lignes,
            AVG(Age) as Age_Moyen,
            NULL as Panier_Moyen,
            NULL as Date_Debut,
            NULL as Date_Fin
        FROM Clients
        """
        
        resume_data = manager.execute_query(resume_query)
        
        # Ajouter la feuille résumé aux données modifiées
        manager.modified_data['Resume'] = resume_data
        
        print("📊 Données du résumé:")
        print(resume_data)
        
        # 4. SAUVEGARDER LE FICHIER MODIFIÉ
        print("\n💾 SAUVEGARDE DU FICHIER MODIFIÉ")
        
        output_filename = f"donnees_modifiees_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        manager.save_excel(output_filename)
        
        print(f"✅ Fichier sauvegardé: {output_filename}")
        
        # 5. VÉRIFICATION FINALE
        print("\n🔍 VÉRIFICATION FINALE")
        
        # Utiliser les données modifiées en mémoire pour les statistiques finales
        ventes_finales = manager.modified_data['Ventes']
        clients_finaux = manager.modified_data['Clients']
        
        ca_final = ventes_finales['Montant_Total'].sum()
        nb_ventes = len(ventes_finales)
        nb_clients = len(clients_finaux)
        
        print(f"   • Ventes: {nb_ventes} lignes, CA: {ca_final:.2f}€")
        print(f"   • Clients: {nb_clients} lignes")
        print(f"   • Feuilles créées: {list(manager.modified_data.keys())}")
        
        return output_filename


def exemple_mise_a_jour_technique():
    """
    Exemple : Mise à jour des données techniques avec corrections et enrichissements.
    """
    print("\n" + "=" * 60)
    print("🔧 EXEMPLE DE MISE À JOUR - Données Techniques")
    print("=" * 60)
    
    with ExcelSQLManager('fiches_lemo_extended.xlsx') as manager:
        # Sauvegarde
        print("💾 Création d'une sauvegarde...")
        manager.backup_original()
        
        # Charger les données
        manager.load_sheets_to_sql()
        
        print("📊 État initial:")
        initial_count = manager.execute_query("SELECT COUNT(*) as total FROM Sheet")
        print(f"   • Nombre total de composants: {initial_count['total'].iloc[0]}")
        
        # 1. MISE À JOUR : Ajouter des catégories de poids
        print("\n🔄 MODIFICATION: Ajout de catégories de poids et corrections")
        
        update_technique_query = """
        SELECT 
            Product,
            NumberContacts,
            WireSize,
            Gender,
            Plug,
            Locking,
            JacketOD_min,
            JacketOD_max,
            RatedCurrent,
            Rmax,
            Vtest_cc,
            Vtest_cs,
            ContactRetention,
            MaxConductor,
            MinConductor,
            BucketDia,
            ContactDia,
            ShellStyle,
            HousingMaterial,
            Keying,
            Colour,
            Variant,
            Weight,
            CASE 
                WHEN Weight > 50 THEN 'Très Lourd'
                WHEN Weight > 30 THEN 'Lourd'
                WHEN Weight > 15 THEN 'Moyen'
                ELSE 'Léger'
            END as Categorie_Poids,
            IP,
            Endurance,
            TempRange,
            Humidity,
            Climatical,
            Shielding_10MHz,
            Shielding_1GHz,
            Shock,
            Vibration,
            SaltSpray,
            CASE 
                WHEN Shielding_10MHz > 75 AND Shielding_1GHz > 40 THEN 'Blindage Excellent'
                WHEN Shielding_10MHz > 65 AND Shielding_1GHz > 35 THEN 'Blindage Bon'
                ELSE 'Blindage Standard'
            END as Qualite_Blindage,
            CASE 
                WHEN NumberContacts >= 8 THEN 'Multi-Contact'
                WHEN NumberContacts >= 5 THEN 'Standard'
                ELSE 'Compact'
            END as Type_Connecteur
        FROM Sheet
        """
        
        # Mettre à jour la feuille
        manager.update_sheet_from_query('Sheet', update_technique_query)
        
        # Vérifier les nouvelles catégories
        categories_poids = manager.execute_query("SELECT Categorie_Poids, COUNT(*) as Nombre FROM Sheet GROUP BY Categorie_Poids ORDER BY Nombre DESC")
        print("⚖️ Répartition par catégorie de poids:")
        print(categories_poids)
        
        qualite_blindage = manager.execute_query("SELECT Qualite_Blindage, COUNT(*) as Nombre FROM Sheet GROUP BY Qualite_Blindage ORDER BY Nombre DESC")
        print("\n🛡️ Répartition par qualité de blindage:")
        print(qualite_blindage)
        
        # 2. CRÉER UNE FEUILLE STATISTIQUES
        stats_query = """
        SELECT 
            'Composants' as Categorie,
            COUNT(*) as Nombre,
            AVG(Weight) as Poids_Moyen,
            AVG(Shielding_10MHz) as Blindage_10MHz_Moyen,
            AVG(NumberContacts) as Contacts_Moyen
        FROM Sheet
        
        UNION ALL
        
        SELECT 
            Gender as Categorie,
            COUNT(*) as Nombre,
            AVG(Weight) as Poids_Moyen,
            AVG(Shielding_10MHz) as Blindage_10MHz_Moyen,
            AVG(NumberContacts) as Contacts_Moyen
        FROM Sheet
        GROUP BY Gender
        """
        
        stats_data = manager.execute_query(stats_query)
        manager.modified_data['Statistiques'] = stats_data
        
        print("\n📈 Statistiques générées:")
        print(stats_data)
        
        # 3. SAUVEGARDER
        output_filename = f"composants_modifies_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        manager.save_excel(output_filename)
        
        print(f"\n✅ Fichier technique sauvegardé: {output_filename}")
        return output_filename


def exemple_comparaison_avant_apres():
    """
    Compare les données avant et après modification.
    """
    print("\n" + "=" * 60)
    print("📊 COMPARAISON AVANT/APRÈS")
    print("=" * 60)
    
    # Lire le fichier original
    print("📖 Lecture du fichier original:")
    with ExcelSQLManager('exemple_donnees.xlsx') as manager:
        manager.load_sheets_to_sql()
        original_ventes = manager.execute_query("SELECT SUM(Montant_Total) as CA_Original FROM Ventes")
        original_count = manager.execute_query("SELECT COUNT(*) as Nb_Original FROM Ventes")
        
        print(f"   • CA Original: {original_ventes['CA_Original'].iloc[0]:.2f}€")
        print(f"   • Nombre de ventes: {original_count['Nb_Original'].iloc[0]}")
    
    # Lire le fichier modifié (le plus récent)
    import glob
    fichiers_modifies = glob.glob("donnees_modifiees_*.xlsx")
    if fichiers_modifies:
        fichier_recent = max(fichiers_modifies)
        print(f"\n📖 Lecture du fichier modifié: {fichier_recent}")
        
        with ExcelSQLManager(fichier_recent) as manager:
            manager.load_sheets_to_sql()
            
            # Analyser les modifications
            modified_ventes = manager.execute_query("SELECT SUM(Montant_Total) as CA_Modifie FROM Ventes")
            remises = manager.execute_query("SELECT COUNT(*) as Nb_Remises FROM Ventes WHERE Statut_Prix = 'Remise 10% appliquée'")
            feuilles = manager.list_tables()
            
            print(f"   • CA Modifié: {modified_ventes['CA_Modifie'].iloc[0]:.2f}€")
            print(f"   • Remises appliquées: {remises['Nb_Remises'].iloc[0]}")
            print(f"   • Feuilles disponibles: {feuilles}")
            
            # Calculer les économies
            economie = original_ventes['CA_Original'].iloc[0] - modified_ventes['CA_Modifie'].iloc[0]
            print(f"\n💰 Économies réalisées avec les remises: {economie:.2f}€")


def main():
    """
    Fonction principale qui exécute tous les exemples de mise à jour.
    """
    print("🔄 EXEMPLES DE MISE À JOUR DE FICHIERS EXCEL")
    print("=" * 70)
    
    try:
        # Exemple 1: Mise à jour des ventes
        fichier_ventes = exemple_mise_a_jour_ventes()
        
        # Exemple 2: Mise à jour technique
        fichier_technique = exemple_mise_a_jour_technique()
        
        # Exemple 3: Comparaison
        exemple_comparaison_avant_apres()
        
        print("\n" + "=" * 70)
        print("✅ TOUS LES EXEMPLES DE MISE À JOUR TERMINÉS !")
        print("=" * 70)
        print(f"📁 Fichiers créés:")
        print(f"   • {fichier_ventes}")
        print(f"   • {fichier_technique}")
        print(f"   • Sauvegardes automatiques créées")
        
        print("\n💡 RÉSUMÉ DES MODIFICATIONS EFFECTUÉES:")
        print("   ✅ Remises automatiques appliquées")
        print("   ✅ Catégories clients ajoutées")
        print("   ✅ Classifications techniques créées")
        print("   ✅ Feuilles de résumé générées")
        print("   ✅ Sauvegardes de sécurité créées")
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()