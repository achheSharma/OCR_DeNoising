import os
from glob import glob
import xmlParser

for file in glob("PDFs/" + "*.pdf"):
    print(file)
    file = file.split("\\")[-1]
    dir = file[:-4]
    print(dir)
    try:
        os.makedirs("XML/" + dir)
    except:
        pass
    os.system("pdftohtml -c -hidden -xml " + "PDFs/" + file + " " + "XML/" + dir + "/" + file + ".xml")
    path = "XML/" + dir + "/"
    xmlParser.parseXML(path, file + ".xml")