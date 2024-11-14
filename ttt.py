import pandas as pd
import glob
import os

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

# Chemin du répertoire contenant les fichiers Excel
repertoire = ''

# Lister tous les fichiers Excel dans le répertoire
fichiers_excel = glob.glob(os.path.join(repertoire, '*.xlsx'))

# Parcourir chaque fichier Excel
for fichier in fichiers_excel:
    # Charger le fichier Excel
    xls = pd.ExcelFile(fichier)
    
    # Parcourir chaque feuille du fichier
    for nom_feuille in xls.sheet_names:
        # Lire la feuille en tant que DataFrame
        df = pd.read_excel(fichier, sheet_name=nom_feuille, dtype=str)
        
        # Vérifier si les colonnes nécessaires existent
        required_columns = ['NO_SIN', 'CD_REFER', 'REF_LB']
        if all(column in df.columns for column in required_columns):
            # Appliquer la fonction pour calculer la colonne 'CLE'
            df['CLE'] = df.apply(calculer_cle, axis=1)
            
            # Définir le nom du nouveau fichier de sortie
            nom_fichier_sortie = f"{os.path.splitext(fichier)[0]}_{nom_feuille}_avec_cle.xlsx"
            
            # Enregistrer la feuille modifiée dans un nouveau fichier Excel
            with pd.ExcelWriter(nom_fichier_sortie, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=nom_feuille, index=False)
            
            print(f"Fichier '{nom_fichier_sortie}' enregistré avec succès.")
        else:
            print(f"Les colonnes nécessaires sont manquantes dans la feuille '{nom_feuille}' du fichier '{fichier}'.")
