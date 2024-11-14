import pandas as pd

# Chemin du fichier Excel
fichier_excel = 'LISTE_VIREMENTS.xlsx'

# Charger le fichier Excel en spécifiant la feuille
df = pd.read_excel(fichier_excel, sheet_name='XT_BASE_VIREMENTS_2024_00', dtype=str)  # Read columns as strings

# Affiche les premières lignes pour vérifier la structure des données
print("Aperçu des données chargées :")
print(df.head())

# Fonction pour calculer la clé selon les conditions
def calculer_cle(row):
    if str(row['NO_SIN']).endswith('73') and row['NO_SIN'] != '000000000000':
        return row['NO_SIN']
    elif str(row['CD_REFER']).endswith('73') and row['CD_REFER'] != '000000000000':
        return row['CD_REFER']
    elif str(row['REF_LB']).endswith('73') and row['REF_LB'] != '000000000000':
        return row['REF_LB']
    else:
        return row['NO_SIN']

# Application de la fonction sur chaque ligne du DataFrame
df['CLE'] = df.apply(calculer_cle, axis=1)

# Affichage du DataFrame mis à jour
print("\nDonnées avec la colonne 'CLE' calculée :")
print(df)

# Enregistrer le DataFrame modifié dans un nouveau fichier Excel
df.to_excel('resultat_avec_cle.xlsx', index=False)
print("\nFichier avec la colonne 'CLE' mis à jour et enregistré sous le nom 'resultat_avec_cle.xlsx'")
