# -*- coding: utf-8 -*-

from datetime import datetime
import pandas as pd

#récupération du mois en cours pour l'import et création du lien image
now=datetime.now()
month=now.strftime("%m")

#Récupération des données du fichier Excel dans un dataframe
x1= pd.read_excel('articles.xlsx',encoding='utf-8') 
pd.set_option('max_colwidth',1000)
se = x1[['Designation','Code','Editeur','Nbre Illustrations','Langue','Prix Vente TTC','Descriptif','Sujet']]
 
se=se.rename(columns = {'Nbre Illustrations':'NbreIllustrations'})
se=se.rename(columns = {'Prix Vente TTC':'PrixVenteTTC'})

#Insertion colonnes supplémentaires
se.insert(1,'Images',0,False)
se.insert(4,'ISBN',0,False)
se.insert(9,'Post_id',0,False)

    
# Test de valeurs non nulle dans le dataframe 
if se.isnull().any().any() == True:    
    print ("Il y a une valeur True dans le tableau donc une case vide")
    print (se.isnull().any) 
    
      
#Modification du contenu des cellules pour obtenir notre code XML
se.loc[:,'Designation']="<item><title>"+se['Designation']+"</title>" 
se.loc[:,'Images']='<content:encoded><![CDATA[<img class="size-full alignleft" src="https://test.files.wordpress.com/2016/'+month+'/'+se['Code']+'.jpg" height="280"/>'
se.loc[:,'Editeur']="Editeur : "+se['Editeur']+"<br/>"
se.loc[:,'NbreIllustrations']="Nombre d'illustrations : "+se['NbreIllustrations'].astype(str)+"</br>"
se.loc[:,'ISBN']="ISBN : "+se['Code']+"</br>"
se.loc[:,'Langue']="Langue : "+se['Langue']+"</br>"
se.loc[:,'PrixVenteTTC']="Prix : "+se['PrixVenteTTC'].astype(str)+" euros</br>"
se.loc[:,'Descriptif']="<p>"+se['Descriptif'].astype(str)+"</p>]]></content:encoded><excerpt:encoded><![CDATA[]]></excerpt:encoded>"
se.loc[:,'Post_id']="<wp:post_id>"+se['Code']+"</wp:post_id>"+"<wp:comment_status>closed</wp:comment_status><wp:status>publish</wp:status><wp:post_type>post</wp:post_type>"
se.loc[:,'Sujet']='<category domain="category"  nicename="'+se['Sujet']+'"><![CDATA['+se['Sujet']+"]]></category></item>"

# suppression de la premiere ligne doublon de l'index extrait du logiciel métier
df = se.ix[1:]

#Conversion du dataframe en string
resultat=df.to_string(header=False,index=False)


#Début du fichier XML obtenu après un export du blog avec un article test
begin="""<?xml version="1.0" encoding="UTF-8"?>
<!-- generator="WordPress.com" created="2015-07-03 12:43"-->
<rss version="2.0" xmlns:excerpt="http://wordpress.org/export/1.2/excerpt/" xmlns:content="http://purl.org/rss/1.0/modules/content/" xmlns:wfw="http://wellformedweb.org/CommentAPI/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:wp="http://wordpress.org/export/1.2/">
  <channel>
<title>Test</title>
<link>https://test.wordpress.com</link>
<description>test</description>
<pubDate>Fri, 03 Jul 2015 12:43:09 +0000</pubDate>
<language>fr</language>
<wp:wxr_version>1.2</wp:wxr_version>
<wp:base_site_url>http://wordpress.com/</wp:base_site_url>
<wp:base_blog_url>https://test.wordpress.com</wp:base_blog_url>
<wp:author>
  <wp:author_login>logintest</wp:author_login>
  <wp:author_email>logintest@test.fr</wp:author_email>
  <wp:author_display_name><![CDATA[User1]]></wp:author_display_name>
  <wp:author_first_name><![CDATA[]]></wp:author_first_name>
  <wp:author_last_name><![CDATA[]]></wp:author_last_name>
</wp:author>
<wp:term>
  <wp:term_id>8119</wp:term_id>
  <wp:term_taxonomy>nav_menu</wp:term_taxonomy>
  <wp:term_slug>menu-1</wp:term_slug>
  <wp:term_name><![CDATA[Menu 1]]></wp:term_name>
</wp:term>
<wp:term>
  <wp:term_id>8533</wp:term_id>
  <wp:term_taxonomy>nav_menu</wp:term_taxonomy>
  <wp:term_slug>menu2</wp:term_slug>
  <wp:term_name><![CDATA[menu2]]></wp:term_name>
</wp:term>
<generator>http://wordpress.com/</generator>
<image>
		<url>https://s2.wp.com/i/buttonw-com.png</url>
		<title>test</title>
		<link>https://test.wordpress.com</link>
</image>"""


#String complete du fichier XML rajoute nos données au début du fichier stantard et on rajoute
# les deux balises de fin channel et rss fermantes.
resultat=begin+resultat+"</channel></rss>"

#Conversion caractère & en code html et encodage en utf-8
donneesClean=resultat.replace("&","&amp;")
donneesClean.encode("utf-8")

# Enregistrement dans fichier xml
mon_fichier = open("fichier.xml", "w") # w option écrase tout 
mon_fichier.write(donneesClean)
mon_fichier.close()

 
   
