from PyPDF2 import PdfFileReader,PdfFileWriter
import os   


#Spliting pdf in two files: *_befor and *_after 
def split(path,name):
    pdf = PdfFileReader(path)
    pdf_writer1 = PdfFileWriter()
    pdf_writer2 = PdfFileWriter()
    for page in range(2):
        pdf_writer1.addPage(pdf.getPage(page))

    result = f'{name}_befor.pdf'
    with open(result,'wb') as output:
        pdf_writer1.write(output)
    
    for page in range(2,pdf.getNumPages()):
        pdf_writer2.addPage(pdf.getPage(page))

    result = f'{name}_after.pdf'
    with open(result,'wb') as output:
        pdf_writer2.write(output)

def merge(path, name):
    df
#Main files directory
MainDir = "C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/split_merge/"
#Secondary files to push into main documents directory
SecondDir = ""
#List of main documents
MainList = []

#Listing all main documents
for root, dirs,files in os.walk(MainDir):
    for file in files:
        if file.endswith(".pdf"):
            MainList.append(file)

for file in MainList:
    split(MainDir + file, file[:file.index('.pdf')])


