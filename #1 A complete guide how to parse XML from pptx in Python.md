# Pptx-Python-guide


# #1 A complete guide how to parse XML from pptx in Python


All module we are needed
```
pip3 install os
pip3 install python-pptx
pip3 install ElementTree
```

# Step 1:

Create your pptx template

# Step 2:

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


# Step 3:


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
        #sql = sql.replace("]]>", "").replace("< ! [CDATA [", "")
        print('SQL:  ',sql)

        
print("-------------------------------------------------------")       
for parametre in root.iter('parametre'):
    print('paramètre: ', parametre.text)

```
