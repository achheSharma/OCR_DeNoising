"""
Main Python script for docx conversion from given PDFs in
specified directory to word documents.

It solves the problem of OCR Denoising when data is to be presented
in word document from flattened scanned PDF files. 

This scripts accepts PDF documents in given specified directory namely 'PDFs/'

It requires poppler utility library to be installed on your machine,
create_docx script which is further dependent on other modules
like docx, spellchecker. Refer README.md for installation instruction.

This script contains following functions:

    * execute - convert PDF file to docx with intermediate parsed xml.
    * main - uses multiprocessing module to execute script for OCR denoising. 
"""

import os
from glob import glob
import create_docx
import multiprocessing as mp
import time

def execute(file):
    """
    Creates intermediate informative xml files for docx conversion to be parsed.
    
    It creates file specific directory inside a XML directory that contains resultant xml and assets of that xml.
    
    Parameters
    ----------
    file : str
        Name of file which will be converted to xml
    """
    # creates directory with same name as xml file.
    dir = file[:-4]
    os.system("pdftohtml -c -hidden -xml " + "PDFs/" + file + " " + "XML/" + dir + "/" + file + ".xml")
    path = "XML/" + dir + "/"
    # call to create_docx script for docx conversion format for output.
    create_docx.create_docx(path, file + ".xml")

if __name__ == '__main__':
    """
    Main function of script that feed all the PDF files with multiprocessing
    module to all the free processors for simultaneous conversion into docx.
    """
    process = []
    PDFs = glob("PDFs/" + "*.pdf")
    totalPDFs = len(PDFs)
    # all the PDFs are passed at once, management and queueing are left onto OS
    # for efficient distribution amongst cores.
    
    # support for batch processing for PDFs can be added
    # with changing number of max_process sent for processing.
    max_process = totalPDFs
    # logic iterating over all PDFs an passing them to execute function
    # for OCR denoising process.
    for i, file in enumerate(PDFs):
        i=i+1
        print(i, ": ", file)
        file = file.split("/")[-1]    # Use backward slash for linux OR '\\' for windows for retrieving path details
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
            # added optional snippet for processing remaining PDFs
            # If batch processing is used for multiprocessing the OCR Denoising Task.
            process = []
    for j in range(len(process)):
        process[j].start()
    for j in range(len(process)):
        process[j].join()
        
        
