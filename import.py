# -*- coding: utf-8 -*-

# Import module normalisation utf8
import unicodedata as ucd

from datetime import datetime
import pandas as pd
x1 = pd.read_excel('articlesOctave.xlsx', encoding='utf-8')

pd.set_option('max_colwidth', 1000)

# Récupération du mois en cours pour l'import et création du lien image
now = datetime.now()
month = now.strftime("%m")

se = x1[['Designation', 'Code', 'Editeur', 'Nbre Illustrations', 'Langue', 'Prix Vente TTC', 'Descriptif', 'Sujet']]

# Normalisation du nom des colonnes
se = se.rename(columns={'Nbre Illustrations': 'NbreIllustrations'})
se = se.rename(columns={'Prix Vente TTC': 'PrixVenteTTC'})

# Insertion colonnes supplémentaires
se.insert(1, 'Images', 0, False)
se.insert(4, 'ISBN', 0, False)
se.insert(9, 'Post_id', 0, False)

# s'il y a des valeurs nulle remplissage avec la valeur inconnu
se = se.fillna('inconnu')

# Test de valeurs non nulle dans le dataframe
if se.isnull().any().any() == True:
    print ("Il y a une valeur True dans le tableau donc une case vide")
    print (se.isnull().any)

# Application normalisation sur les colonnes contenant du texte
se['Designation'] = se['Designation'].map(lambda x: ucd.normalize('NFKD', x))
se['Descriptif'] = se['Descriptif'].map(lambda x: ucd.normalize('NFKD', x))

# Modification des données
se.loc[:, 'Designation'] = "<item><title>"+se['Designation']+"</title>"
se.loc[:, 'Images'] = '<content:encoded><![CDATA[<img class="size-full alignleft" src="https://interartparis.files.wordpress.com/2016/'+month+'/'+se['Code']+'.jpg" height="280"/>'
se.loc[:, 'Editeur'] = "Editeur : "+se['Editeur']+"<br/>"
se.loc[:, 'NbreIllustrations'] = "Nombre d'illustrations : "+se['NbreIllustrations'].astype(str)+"</br>"
se.loc[:, 'ISBN'] = "ISBN : "+se['Code']+"</br>"
se.loc[:, 'Langue'] = "Langue : "+se['Langue']+"</br>"
se.loc[:, 'PrixVenteTTC'] = "Prix : "+se['PrixVenteTTC'].astype(str)+" euros</br>"
se.loc[:, 'Descriptif'] = "<p>"+se['Descriptif'].astype(str)+"</p>]]></content:encoded><excerpt:encoded><![CDATA[]]></excerpt:encoded>"
se.loc[:, 'Post_id'] = "<wp:post_id>"+se['Code']+"</wp:post_id>"+"<wp:comment_status>closed</wp:comment_status><wp:status>publish</wp:status><wp:post_type>post</wp:post_type>"
se.loc[:, 'Sujet'] = '<category domain="category"  nicename="'+se['Sujet']+'"><![CDATA['+se['Sujet']+"]]></category></item>"


# Suppression de la colonne Code
del se['Code']

# Suppression de la première ligne doublon de l'index extrait du logiciel métier
df = se.ix[1:]

# Conversion du dataframe en string
resultat = df.to_string(header=False, index=False)

# Ajout des balises de fin en xml
resultat = resultat+"</channel></rss>"

# Conversion caractère & en code html et encodage en utf-8
donneesClean = resultat.replace("&", "&amp;")
donneesClean.encode("utf-8")

# Enregistrement dans fichier xml
mon_fichier = open('fichier.xml', encoding='utf-8', mode='a')  # option a append
mon_fichier.write(donneesClean)
mon_fichier.close()

print("Well done")

