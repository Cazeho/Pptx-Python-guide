Pptx-Python-guide#pptx-python-guide

# Pptx-Python-guide


# #1 A complete guide how to parse XML from pptx in Python

All the modules we need
```
pip3 install os-sys
pip3 install python-pptx
pip3 install lxml
pip3 install elementpath
pip3 install datetime
```

# Step 1: Create your template

Create your pptx template: template.pptx

![template](https://user-images.githubusercontent.com/58745332/83173894-d9e2df80-a119-11ea-9b54-3f0452b65726.PNG)

# Step 2: Convert pptx to XML

templateConfig.py

```
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET
import os 

pptxTemplate = "template.pptx"
templateConfig = "templateConfig.xml"
prs = Presentation(pptxTemplate)
presentationLevel = ET.Element("presentation",file_name= pptxTemplate, author="....")
slidesLevel = ET.SubElement(presentationLevel, "slides")
num = 1
#On parcourt les éléments de pw
for slide in prs.slides:
	slideLevel = ET.SubElement(slidesLevel, 'slide',slide_num= str(num),slide_id= str(slide.slide_id))
	for shape in slide.shapes:
		shapeLevel = ET.SubElement(slideLevel, 'shape',shape_id= str(shape.shape_id), shape_name= shape.name,shape_type=str(shape.shape_type), left= str(round(shape.left*100/12204000,0)), top= str(round(shape.top*100/6840000,0)), width= str(round(shape.width*100/12204000,0)), height= str(round(shape.height*100/6840000,0)))
		sqlLevel = ET.SubElement(shapeLevel, 'sql')
	num += 1
parametresLevel = ET.SubElement(presentationLevel, "parametres")
parametreLevel = ET.SubElement(parametresLevel, "parametre",parametre_name="",parametre_type="")

tree = ET.ElementTree(presentationLevel)
tree.write(templateConfig)
```

templateConfig.xml

```
<?xml version="1.0"?>
<presentation file_name="template.pptx" author="....">
   <slides>
       <slide slide_num="1" slide_id="256">
           <shape width="6.0" top="12.0" shape_type="TEXT_BOX (17)" shape_name="ZoneTexte 1" shape_id="2" left="9.0" height="5.0">
              <sql>< ! [CDATA [lolo]]></sql>
           </shape>
           <shape width="40.0" top="25.0" shape_type="PICTURE (13)" shape_name="Image 3" shape_id="4" left="27.0" height="55.0">
              <sql>< ! [CDATA [lolo]]></sql>
           </shape>
       </slide>
       <slide slide_num="2" slide_id="257">
           <shape width="26.0" top="15.0" shape_type="TEXT_BOX (17)" shape_name="ZoneTexte 1" shape_id="2" left="10.0" height="5.0">
              <sql>< ! [CDATA [lolo]]></sql>
	   </shape>
	   <shape width="59.0" top="21.0" shape_type="CHART (3)" shape_name="Graphique 4" shape_id="5" left="20.0" height="59.0">
	      <sql>< ! [CDATA [lolo]]></sql>
	   </shape>
       </slide>
       <slide slide_num="3" slide_id="258">
           <shape width="6.0" top="13.0" shape_type="TEXT_BOX (17)" shape_name="ZoneTexte 1" shape_id="2" left="7.0" height="5.0">
              <sql>< ! [CDATA [lolo]]></sql>
	   </shape>
           <shape width="67.0" top="34.0" shape_type="TABLE (19)" shape_name="Tableau 3" shape_id="4" left="17.0" height="33.0">
              <sql>< ! [CDATA [lolo]]></sql>
	   </shape>
       </slide>
   </slides>
   <parametres>
      <parametre parametre_type="" parametre_name=""/>
   </parametres>
</presentation>
```


# Step 3: Parse XML with Python


parse.py



```
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import parse
from lxml import etree


root = ET.parse('templateConfig.xml').getroot()
tag = root.tag
print(tag) 
attributes = root.attrib
print(attributes)



for slide in root.iter("slide"):
    print("-------------------------------------------------------")
    print ('slide :',slide.attrib["slide_num"] + " " + 'slide_id :',slide.attrib["slide_id"])
    for shape in slide.iter("shape"):
        print('shape_id:', shape.attrib["shape_id"] + " " + 'shape_type:',shape.attrib["shape_type"] + " " + 'shape_name:', shape.attrib["shape_name"])
        sql = shape.find('sql').text
        sql = sql.replace("]]>", "").replace("< ! [CDATA [", "")
        print('SQL:  ',sql)

        
print("-------------------------------------------------------")       
for parametre in root.iter('parametre'):
    print('paramètre: ', parametre.text)

```

# Output

```
presentation
{'author': '....', 'file_name': 'template.pptx'}
-------------------------------------------------------
slide : 1 slide_id : 256
shape_id: 2 shape_type: TEXT_BOX (17) shape_name: Text 1
SQL:   [$text1]
shape_id: 4 shape_type: PICTURE (13) shape_name: picture 3
SQL:   [$image_bin]
-------------------------------------------------------
slide : 2 slide_id : 257
shape_id: 2 shape_type: TEXT_BOX (17) shape_name: Text 1
SQL:   [$text2]
shape_id: 5 shape_type: CHART (3) shape_name: Chart 4
SQL:   [$chart]
-------------------------------------------------------
slide : 3 slide_id : 258
shape_id: 2 shape_type: TEXT_BOX (17) shape_name: Text 1
SQL:   [$texte3]
shape_id: 4 shape_type: TABLE (19) shape_name: Table 3
SQL:   [$tableau]
-------------------------------------------------------
paramètre:  None
paramètre:  [$text1]
paramètre:  [$image_bin]
paramètre:  [$text2]
paramètre:  [$chart]
paramètre:  [$texte3]
paramètre:  [$tableau]

```


`code made in France 2020`
