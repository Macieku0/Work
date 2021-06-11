from PyPDF2 import PdfFileReader,PdfFileWriter
import os   

#Needs to be open in admin cmd
#Required three folders named "Main", "Secondary" and "Split"
#Main folder is for main documents befor and after script work
#Split folder is for splited documents
#Secondary folder is for documents that needs to be inserted between splited documents



#Spliting pdf in two files: *_befor and *_after 
def split(path,name,dir):
    pdf = PdfFileReader(path)
    pdf_writer1 = PdfFileWriter()
    pdf_writer2 = PdfFileWriter()
    for page in range(2):
        pdf_writer1.addPage(pdf.getPage(page))

    result = os.path.join(dir,f'{name}_befor.pdf')
    with open(result,'wb') as output:
        pdf_writer1.write(output)
    
    for page in range(2,pdf.getNumPages()):
        pdf_writer2.addPage(pdf.getPage(page))

    result = os.path.join(dir,f'{name}_after.pdf')
    with open(result,'wb') as output:
        pdf_writer2.write(output)
#Merging all files 
def merge(paths,name,dir):
    pdf_writer = PdfFileWriter()
    for path in paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))
    
    result = os.path.join(dir,f'{name}.pdf')
    with open(result,'wb') as output:
        pdf_writer.write(output)

if __name__ == '__main__':
    #Secondary files have to be named in specific format = "{Main File_name}_between.pdf"
    Directory  = "C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/split_merge"
    #Main files directory
    MainDir = Directory + "/Main/"
    #Where should splitted files be stored
    SplitDir = Directory + "/Split/"
    #Secondary files to push into main documents directory
    SecondDir = Directory + "/Secondary/"
    #List of main documents
    MainList = []




    #Listing all main documents
    for root, dirs,files in os.walk(MainDir):
        for file in files:
            if file.endswith(".pdf"):
                MainList.append(file)

    for file in MainList:
        split(MainDir + file, file[:file.index('.pdf')],SplitDir)
        list = [f"{SplitDir}{file[:file.index('.pdf')]}_befor.pdf",f"{SecondDir}{file[:file.index('.pdf')]}_between.pdf",f"{SplitDir}{file[:file.index('.pdf')]}_after.pdf"]
        merge(list,file[:file.index('.pdf')],MainDir)

