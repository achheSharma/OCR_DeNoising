import xml.etree.ElementTree as ET
from pprintpp import pprint as pp # pretty print karane ke liye dictionary ko
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import re
import copy

fonts = {}
prevTop = None
fontStyle = [False, False, False] # Bold, Italics, Underline
def DFS(parent, page_data, fontStyle):
    for elem in parent:
        if elem.tag == 'fontspec':
            fonts[elem.attrib['id']] = elem.attrib
            continue

        # if elem.tag == 'b':
        #     return elem.tag, elem.text
        text, fontStyle = DFS(elem, page_data, copy.deepcopy(fontStyle))
        if elem.tag == 'b':
            fontStyle[0] = True
            return elem.text, fontStyle
        elif elem.tag == 'i':
            fontStyle[1] = True
            return elem.text, fontStyle
        elif elem.tag == 'u':
            fontStyle[2] = True
            return elem.text, fontStyle

        if elem.tag == 'text':
            if not text:
                text = elem.text
            if not text or text == ' ':
                continue
            attrib = elem.attrib
            attrib['b'] = fontStyle[0]
            attrib['i'] = fontStyle[1]
            attrib['u'] = fontStyle[2]
            if elem.tag in page_data:
                page_data[elem.tag].append([text, attrib])
            else:
                page_data[elem.tag] = [[text, attrib]]
    return None, fontStyle

xml = "XML\\A7H74ZZ6\\A7H74ZZ6.pdf.xml"
tree  = ET.parse(xml)
root = tree.getroot()

pages = []
pageHeight = None
pageWidth = None
for page in root:
    page_data = {}
    fontStyle = [False, False, False]
    DFS(page, page_data, fontStyle)
    pages.append(page_data)
    pageHeight = int(page.attrib['height'])
    pageWidth = int(page.attrib['width'])
    # break
# pp(pages)

# check image addition code ???
# https://stackoverflow.com/questions/32932230/add-an-image-in-a-specific-position-in-the-document-docx-with-python
# loop structure inconsistent

def style(run, text, attrib, prevLeft = None):
    # print(text)
    font = run.font
    tab = ""
    if prevLeft:
        diff = abs(int(attrib['left']) - prevLeft)
        if diff > 50:
            # print("\tTAB: ", diff)
            for i in range(int(diff/50)):
                tab = tab + "\t"
    run.text = tab + text + " "
    font.name = fonts[attrib['font']]['family']
    font.size = Pt(int(fonts[attrib['font']]['size']))
    if 'b' in attrib.keys():
        font.bold = attrib['b']
    if 'i' in attrib.keys():
        font.italic = attrib['i']
    if 'u' in attrib.keys():
        font.underline = attrib['u']

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
            prevLeft = None
            prevTop = None
            for item in value:
                text = item[0]
                attrib = item[1]
                if top != attrib['top']:
                    top = attrib['top']
                    left = int(attrib['left'])
                    para = document.add_paragraph()
                    para_format = para.paragraph_format
                    if prevTop:
                        if abs(int(top) - prevTop) < 20:
                            para_format.space_after = Pt(0)
                    run = para.add_run()
                    style(run, text, attrib)
                else:
                    run = para.add_run()
                    style(run, text, attrib, prevLeft)
                if left > 200 and abs(pageWidth - (left + int(attrib['width'])) - int(attrib['left'])) < 25:
                        para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # else:
                #     para_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                prevLeft = int(attrib['left']) + int(attrib['width'])
                prevTop = int(attrib['top']) + int(attrib['height'])
        
        # if key == 'image':
        #     for item in value:
        #         attrib = item[1]
        #         img = attrib['src']
        #         mtch = re.match(r'_1.png',img) # avoid first image of complete page
        #         if(img):
        #             document.add_picture(img)            
    
    document.add_page_break()

document.save('demo_A7H74ZZ6.docx')
# check afterwards
