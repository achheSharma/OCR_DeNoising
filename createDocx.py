import xml.etree.ElementTree as ET
from pprintpp import pprint as pp # pretty print karane ke liye dictionary ko
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import re

fonts = {}
prevTop = None
def DFS(parent, page_data):
    for elem in parent:
        if elem.tag == 'fontspec':
            fonts[elem.attrib['id']] = elem.attrib
            continue

        if elem.tag == 'b':
            return elem.tag, elem.text
        tag, text = DFS(elem, page_data)
        attrib = elem.attrib
        if tag == 'b':
            attrib[tag] = True
        else:
            text = elem.text        
        if elem.tag in page_data:
            page_data[elem.tag].append([text, attrib])
        else:
            page_data[elem.tag] = [[text, attrib]]
    return None, None

xml = "XML\\A7FX7D8H\\new_A7FX7D8H.pdf.xml"
tree  = ET.parse(xml)
root = tree.getroot()

pages = []
pageHeight = None
pageWidth = None
for page in root:
    page_data = {}
    DFS(page, page_data)
    pages.append(page_data)
    pageHeight = int(page.attrib['height'])
    pageWidth = int(page.attrib['width'])
    # break
# pp(pages)

# check image addition code ???
# https://stackoverflow.com/questions/32932230/add-an-image-in-a-specific-position-in-the-document-docx-with-python
# loop structure inconsistent

def style(run, text, attrib):
    font = run.font
    run.text = text + " "
    font.name = fonts[attrib['font']]['family']
    font.size = Pt(int(fonts[attrib['font']]['size']))
    if 'b' in attrib.keys():
        font.bold = attrib['b']

# pp(pages)

document = Document()

for page in pages:
    top = None
    for key, value in page.items():
        if key == 'text':
            value.sort(key = lambda x: int(x[1]['top']))
            for i in range(len(value)-1):
                if abs(int(value[i][1]['top']) - int(value[i+1][1]['top'])) <= 5:
                    value[i+1][1]['top'] = value[i][1]['top']
            value.sort(key = lambda x: (int(x[1]['top']), int(x[1]['left'])))
            # pp(value)
            for item in value:
                text = item[0]
                attrib = item[1]
                # if top == None or abs(int(top) - int(attrib['top'])) > 10:
                if top != attrib['top']:
                    top = attrib['top']
                    left = attrib['left']
                    para = document.add_paragraph()
                    para_format = para.paragraph_format
                    run = para.add_run()
                    style(run, text, attrib)
                else:
                    run = para.add_run()
                    style(run, text, attrib)

                if int(left) > 200 and abs(pageWidth - (int(left) + int(attrib['width'])) - int(attrib['left'])) < 25:
                        para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    para_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # if key == 'image':
        #     for item in value:
        #         attrib = item[1]
        #         img = attrib['src']
        #         mtch = re.match(r'_1.png',img) # avoid first image of complete page
        #         if(img):
        #             document.add_picture(img)            
    
    document.add_page_break()

document.save('demo.docx')
# check afterwards
