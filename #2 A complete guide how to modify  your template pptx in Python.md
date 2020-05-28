# Pptx-Python-guide

# #2 A complete guide how to modify  your template pptx in Python

All module we are needed
```
pip3 install os-sys
pip3 install python-pptx
pip3 install lxml
pip3 install elementpath
pip3 install datetime
```

# Step 1 : 

![template](https://user-images.githubusercontent.com/58745332/83173894-d9e2df80-a119-11ea-9b54-3f0452b65726.PNG)

# Step 2 : 

```
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from datetime import date

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import parse
from lxml import etree

import os

def ecrire(sl_id,sh_id,sh_type,d):#fonction pour modifier les différents types de placeholder
    for slide in prs.slides:
        if slide.slide_id == sl_id:
            for shape in slide.shapes:
                if shape.shape_id == sh_id:                  
                    if sh_type == MSO_SHAPE_TYPE.TEXT_BOX : # Si le shape est un texte
                        shape.text = d

                    if sh_type == MSO_SHAPE_TYPE.TABLE : # si le shape est un tableau
                        for l in range(5):
                            for c in range(5):
                                print(c)
                                shape.table.cell(l,c).text = d[l][c]
                    if sh_type == MSO_SHAPE_TYPE.CHART :
                        shape.chart.replace_data(d)
                        print(sh_type)
                    if sh_type == MSO_SHAPE_TYPE.PICTURE : # Si le shape est une image
                        print (d)
                        
prs = Presentation('template.pptx')


#Permet de créer des listes
data_slide=[]
data_shape=[]
data_type=[]


root = ET.parse('exemple.pptx.xml').getroot()#charge le template (si sql décommenter) (1)
#root = ET.parse('templateConfig.xml').getroot()#si sql commenter
tag = root.tag
print(tag) 
attributes = root.attrib
print(attributes)


#Parsing
for slide in root.iter("slide"):#regarde au niveau du slide
    print("-------------------------------------------------------")
    print ('slide :',slide.attrib["slide_num"] + " " + 'slide_id :',slide.attrib["slide_id"])
    for shape in slide.iter("shape"):#regarde au niveau du shape
        a=slide.attrib["slide_id"]
        b= shape.attrib["shape_id"]
        c=shape.attrib["shape_type"].replace("TEXT_BOX (", "").replace("TABLE (", "").replace("CHART (", "").replace("PICTURE (", "").replace(")", "")#supprime les differents labels de type text,tab,chart et image
        data_slide.append(a)#insère les elements dans la liste
        data_shape.append(b)
        data_type.append(c)
        #b= shape.attrib["shape_type"]
        #print (char.attrib, char.text)
        print('shape_id:', shape.attrib["shape_id"] + " " + 'shape_type:',shape.attrib["shape_type"] + " " + 'shape_name:', shape.attrib["shape_name"])#affiche les infos du xml
        sql = shape.find('sql').text
        sql = sql.replace("]]>", "").replace("< ! [CDATA [", "")#primtive xml/supprime le CDATA/(si sql décommenter) (1)
        print('SQL:  ',sql)#Affiche le sql

        
print("-------------------------------------------------------")       
for parametre in root.iter('parametre'):#Affiche les paramètres
    print('paramètre: ', parametre.text)







print(data_slide)#affiche la liste
print(data_shape)
print(data_type)

#print(int(data_slide[0]))
#print(int(data_shape[0]))
#print(int(data_type[0]))








# Préparation des données à afficher dans les tableaux d'au moins 5x5
data_table = [['colonne1','colonne2','colonne3','colonne4','colonne5'],['cellule_1x1','cellule_1x2','cellule_1x3','cellule_1x4','cellule_1x5'],['cell_2x1','cell_2x2','cell_2x3','cell_2x4','cell_2x5'],['','','','',''],['cell_4x1','cell_4x2','cell_4x3','cell_4x4','cell_4x5']]
# Préparation des données alimentant les graphiques
data_chart = CategoryChartData()
data_chart.categories = ['Test1', 'Test1', 'Test3']
data_chart.add_series('Risques', (10.2, 21.4, 16.7))
data_chart.add_series('Vulnérabilités', (19.2, 43.4, 8.7))
data_chart.add_series('Couverture', (13.2, 19.4, 6.7))

# Modifier le shape (id = 4 de type tableau (19) de la slside 258)

data_titre1="TITRE CHART"

data_titre2="TITRE TAB"

# Préparation du texte à afficher dans les textbox
data_textbox = "Rapport généré le {:%d/%m/%Y-%H:%M:%S}".format(date.today())
ecrire(int(data_slide[0]),int(data_shape[0]),int(data_type[0]),data_textbox)


ecrire(int(data_slide[2]),int(data_shape[2]),int(data_type[2]),data_titre1)

ecrire(int(data_slide[4]),int(data_shape[4]),int(data_type[4]),data_titre2)


ecrire(int(data_slide[3]),int(data_shape[3]),int(data_type[3]),data_chart)

ecrire(int(data_slide[5]),int(data_shape[5]),int(data_type[5]),data_table)





prs.save('report.pptx')#save all modifications

os.startfile('report.pptx')#ouvre le pptx juste après l'execution


```

# Output

![rapport](https://user-images.githubusercontent.com/58745332/83176761-4d86eb80-a11e-11ea-9d58-27ad81b50600.PNG)



`code made in France 2020`

