import os
from glob import glob
import createDocx
import multiprocessing as mp
import time

def execute(file):
    dir = file[:-4]
    os.system("pdftohtml -c -hidden -xml " + "PDFs/" + file + " " + "XML/" + dir + "/" + file + ".xml")
    path = "XML/" + dir + "/"
    createDocx.createDocx(path, file + ".xml")

if __name__ == '__main__':
    process = []
    max_process = 4
    PDFs = glob("PDFs/" + "*.pdf")
    totalPDFs = len(PDFs)
    # print(totalPDFs)
    for i, file in enumerate(PDFs):
        i=i+1
        print(i, ": ", file)
        file = file.split("\\")[-1]
        try:
            os.makedirs("XML/" + file[:-4])
        except:
            pass
        
        p = mp.Process(target = execute, args = (file,))
        process.append(p)
        if i%max_process==0:
            for j in range(len(process)):
                process[j].start()
            for j in range(len(process)):
                process[j].join()
            process = []
    for j in range(len(process)):
        process[j].start()
    for j in range(len(process)):
        process[j].join()
    process = []
