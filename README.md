## Technical Components  
Platform Tested  
• Windows 10  
Tools  
• Poppler for Windows  
• Python  
Language Libraries  
• elementtree  
• python-docx  
• pyspellchecker  
  
## Files:  
1. main.py  
• Executes poppler to generate xml of respective pdfs present in PDFs forlder  
• Saves the XML to respective folder in the "XML" folder  
• Calls createDocx.py  
  
2. createDocx.py  
• Parses the previously generated XML file  
• Performs De-Noising on the textual data  
• Created Docx as similar to the original PDF in representation as possible  

3. words.txt  
• Contains list of words to be excluded from spell checking  

## Steps for Execution:   
• Install all pre requisite tools needed  
• Place all PDF documents in PDFs folder  
• Execute main.py  

## Results:  
• Subfolders with same name as that of file are created in the DOCX directory  
• This folder holds the generated docx file for each respective PDF provided
