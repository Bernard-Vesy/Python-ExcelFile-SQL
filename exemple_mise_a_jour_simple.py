"""
Exemple SIMPLE de mise à jour de données Excel avec SQL
Démonstration claire et fonctionnelle
"""

from excel_sql_manager import ExcelSQLManager
import pandas as pd
from pathlib import Path

def exemple_simple_mise_a_jour():
    """
    Exemple simple et complet de mise à jour de données Excel.
    """
    print("🔄 EXEMPLE SIMPLE DE MISE À JOUR DE DONNÉES EXCEL")
    print("=" * 60)
    
    # S'assurer que le fichier d'exemple existe
    if not Path('exemple_donnees.xlsx').exists():
        print("⚠️  Création du fichier d'exemple...")
        from example_usage import creer_fichier_exemple
        creer_fichier_exemple()
    
    with ExcelSQLManager('exemple_donnees.xlsx') as manager:
        # 1. SAUVEGARDE AUTOMATIQUE
        print("💾 Création d'une sauvegarde de sécurité...")
        manager.backup_original()
        print("✅ Sauvegarde créée")
        
        # 2. CHARGEMENT DES DONNÉES
        print("\n📊 Chargement des données originales...")
        manager.load_sheets_to_sql()
        
        # Afficher l'état initial
        initial_data = manager.execute_query("SELECT COUNT(*) as nb_ventes, SUM(Montant_Total) as ca_total FROM Ventes")
        print(f"   • Nombre de ventes: {initial_data['nb_ventes'].iloc[0]}")
        print(f"   • CA initial: {initial_data['ca_total'].iloc[0]:.2f}€")
        
        # 3. PREMIÈRE MODIFICATION : Appliquer des remises
        print("\n🔄 MODIFICATION 1: Application de remises automatiques")
        print("   → Remise de 15% sur les montants > 5000€")
        print("   → Remise de 5% sur les montants entre 3000€ et 5000€")
        
        requete_remises = """
        SELECT 
            Date,
            Produit,
            Quantite,
            Prix_Unitaire,
            Vendeur,
            CASE 
                WHEN Montant_Total > 5000 THEN ROUND(Montant_Total * 0.85, 2)
                WHEN Montant_Total > 3000 THEN ROUND(Montant_Total * 0.95, 2)
                ELSE Montant_Total
            END as Montant_Total,
            CASE 
                WHEN Montant_Total > 5000 THEN 'Remise 15%'
                WHEN Montant_Total > 3000 THEN 'Remise 5%'
                ELSE 'Prix normal'
            END as Type_Remise
        FROM Ventes
        ORDER BY Date
        """
        
        # Appliquer la modification
        manager.update_sheet_from_query('Ventes', requete_remises)
        
        # Vérifier les résultats
        ventes_modifiees = manager.modified_data['Ventes']
        remises_15 = len(ventes_modifiees[ventes_modifiees['Type_Remise'] == 'Remise 15%'])
        remises_5 = len(ventes_modifiees[ventes_modifiees['Type_Remise'] == 'Remise 5%'])
        nouveau_ca = ventes_modifiees['Montant_Total'].sum()
        
        print(f"✅ Remises appliquées:")
        print(f"   • {remises_15} ventes avec remise 15%")
        print(f"   • {remises_5} ventes avec remise 5%")
        print(f"   • Nouveau CA: {nouveau_ca:.2f}€")
        print(f"   • Économie réalisée: {initial_data['ca_total'].iloc[0] - nouveau_ca:.2f}€")
        
        # 4. DEUXIÈME MODIFICATION : Enrichir les données clients
        print("\n🔄 MODIFICATION 2: Enrichissement des données clients")
        print("   → Ajout de catégories d'âge")
        print("   → Classification par segment")
        
        requete_clients = """
        SELECT 
            ID_Client,
            Nom,
            Ville,
            Age,
            Statut,
            CASE 
                WHEN Age >= 65 THEN 'Senior'
                WHEN Age >= 45 THEN 'Adulte'
                WHEN Age >= 25 THEN 'Jeune Adulte' 
                ELSE 'Jeune'
            END as Tranche_Age,
            CASE 
                WHEN Statut = 'Actif' AND Age >= 50 THEN 'VIP'
                WHEN Statut = 'Actif' THEN 'Standard'
                ELSE 'Prospect'
            END as Segment
        FROM Clients
        ORDER BY Age DESC
        """
        
        manager.update_sheet_from_query('Clients', requete_clients)
        
        # Analyser les nouveaux segments
        clients_modifies = manager.modified_data['Clients']
        segments = clients_modifies.groupby(['Tranche_Age', 'Segment']).size().reset_index(name='Nombre')
        print("✅ Nouveaux segments créés:")
        for _, row in segments.iterrows():
            print(f"   • {row['Tranche_Age']} - {row['Segment']}: {row['Nombre']} clients")
        
        # 5. CRÉATION D'UNE FEUILLE DE SYNTHÈSE
        print("\n📋 CRÉATION D'UNE FEUILLE DE SYNTHÈSE")
        
        # Préparer les données de synthèse
        synthese_data = {
            'Indicateur': [
                'Nombre total de ventes',
                'CA après remises',
                'Économies réalisées',
                'Nombre de clients',
                'Clients VIP',
                'Taux de remise moyen'
            ],
            'Valeur': [
                len(ventes_modifiees),
                f"{nouveau_ca:.2f}€",
                f"{initial_data['ca_total'].iloc[0] - nouveau_ca:.2f}€",
                len(clients_modifies),
                len(clients_modifies[clients_modifies['Segment'] == 'VIP']),
                f"{((initial_data['ca_total'].iloc[0] - nouveau_ca) / initial_data['ca_total'].iloc[0] * 100):.1f}%"
            ],
            'Date_Maj': [pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')] * 6
        }
        
        synthese_df = pd.DataFrame(synthese_data)
        manager.modified_data['Synthese'] = synthese_df
        
        print("✅ Feuille de synthèse créée avec les KPIs principaux")
        
        # 6. SAUVEGARDE FINALE
        print("\n💾 SAUVEGARDE DU FICHIER FINAL")
        fichier_final = f"donnees_finales_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        manager.save_excel(fichier_final)
        
        print(f"✅ Fichier sauvegardé: {fichier_final}")
        print(f"📊 Feuilles créées: {list(manager.modified_data.keys())}")
        
        # 7. RÉSUMÉ FINAL
        print("\n" + "=" * 60)
        print("📈 RÉSUMÉ DES MODIFICATIONS RÉALISÉES")
        print("=" * 60)
        print(f"✅ {remises_15 + remises_5} ventes modifiées avec remises")
        print(f"✅ {len(clients_modifies)} clients segmentés")
        print(f"✅ 3 feuilles Excel créées (Ventes, Clients, Synthese)")
        print(f"✅ Économies: {initial_data['ca_total'].iloc[0] - nouveau_ca:.2f}€")
        print(f"✅ Fichier sauvé: {fichier_final}")
        print("=" * 60)
        
        return fichier_final


def demo_verification_fichier(nom_fichier):
    """
    Démontre comment vérifier le contenu du fichier modifié.
    """
    print(f"\n🔍 VÉRIFICATION DU FICHIER CRÉÉ: {nom_fichier}")
    print("=" * 60)
    
    if not Path(nom_fichier).exists():
        print("❌ Fichier non trouvé")
        return
    
    with ExcelSQLManager(nom_fichier) as manager:
        # Charger le fichier modifié
        manager.load_sheets_to_sql()
        
        print("📊 Contenu du fichier modifié:")
        tables = manager.list_tables()
        
        for table in tables:
            info = manager.get_table_info(table)
            print(f"   • {table}: {info['row_count']} lignes")
            
            # Aperçu de chaque feuille
            if table == 'Ventes':
                apercu = manager.execute_query("SELECT Type_Remise, COUNT(*) as nb FROM Ventes GROUP BY Type_Remise")
                print(f"     Répartition des remises:")
                for _, row in apercu.iterrows():
                    print(f"       - {row['Type_Remise']}: {row['nb']}")
                    
            elif table == 'Clients':
                apercu = manager.execute_query("SELECT Segment, COUNT(*) as nb FROM Clients GROUP BY Segment")
                print(f"     Répartition des segments:")
                for _, row in apercu.iterrows():
                    print(f"       - {row['Segment']}: {row['nb']}")
                    
            elif table == 'Synthese':
                apercu = manager.preview_data(table, 10)
                print(f"     Indicateurs clés:")
                for _, row in apercu.iterrows():
                    print(f"       - {row['Indicateur']}: {row['Valeur']}")


def main():
    """
    Fonction principale pour exécuter l'exemple complet.
    """
    print("🚀 DÉMONSTRATION COMPLÈTE DE MISE À JOUR EXCEL")
    print("=" * 70)
    
    try:
        # Exécuter l'exemple principal
        fichier_cree = exemple_simple_mise_a_jour()
        
        # Vérifier le résultat
        demo_verification_fichier(fichier_cree)
        
        print("\n🎉 DÉMONSTRATION TERMINÉE AVEC SUCCÈS !")
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()