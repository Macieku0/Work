from PyPDF2 import PdfFileReader,PdfFileWriter
import os   

#Needs to be open in admin cmd
#Required three folders named "Main", "Secondary" and "Split"
#Main folder is for main documents befor and after script work
#Split folder is for splited documents
#Secondary folder is for documents that needs to be inserted between splited documents



#Spliting pdf in two files: *_befor and *_after 
def split(path,name,dir,length):
    pdf = PdfFileReader(path)
    pdf_writer1 = PdfFileWriter()
    pdf_writer2 = PdfFileWriter()
    for page in range(7):
        pdf_writer1.addPage(pdf.getPage(page))

    result = os.path.join(dir,f'{name}_befor.pdf')
    with open(result,'wb') as output:
        pdf_writer1.write(output)
    
    for page in range(7+length,pdf.getNumPages()):
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


def countPages(a,b):
    length = 0
    filesList = []
    for root, dirs,files in os.walk(a):
        for file in files:
            if file != b:
                pdf = PdfFileReader(f'{a}{file}')
                length += pdf.getNumPages()
                filesList.append(file)
    return length, filesList

if __name__ == '__main__':
    #Secondary files have to be named in specific format = "{Main File_name}_between.pdf"
    Directory  = "C:\\Users\\macie\\Pulpit\\III_CAT_SPRAWDZENIE_SRUB\\Do przedruku\\"
    #Main files directory
    MainDir = Directory #+ "/Main/"
    #Where should splitted files be stored
    SplitDir = Directory + "Split\\"
    #Secondary files to push into main documents directory
    SecondDir = Directory + "/Secondary/"
    #List of main documents
    MainList = []




    #Listing all main documents
    # for root, dirs,files in os.walk(MainDir):
    #     for file in files:
    #         if file.endswith(".pdf"):
    #             MainList.append(file)
    MainList = ['200-ANS72-3311070-EC07CPC-S60.pdf']
    for file in MainList:
        MainDir = f'{Directory}{file[:file.index(".pdf")]}\\'
        length, secondList = countPages(MainDir,file)
        split(MainDir + file, file[:file.index('.pdf')],SplitDir,length)
        list = [f"{SplitDir}{file[:file.index('.pdf')]}_befor.pdf"]
        for files in secondList:
            list.append(f'{MainDir}{files[:files.index(".pdf")]}.pdf')
        list.append(f"{SplitDir}{file[:file.index('.pdf')]}_after.pdf")
        merge(list,file[:file.index('.pdf')],MainDir)

