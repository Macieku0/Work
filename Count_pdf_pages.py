from numpy import roots
import pandas as pd
import os
from PyPDF2 import PdfFileReader

pdf_pages = 0

#Directory with pdf files
dirPathTo = 'C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/PDF/'

#Report structure
df = pd.DataFrame(columns=['fileName', 'fileLocation', 'pageNumber'])

#Directory to save report
Report = open(dirPathTo + '/' + 'Count_pages_report.csv', 'w')
#Main loop
for root, dirs, files in os.walk(dirPathTo):
    for f in files:
        if f.endswith('.pdf'):
            #Read each pdf in directory
            pdf = PdfFileReader(open(os.path.join(root, f), 'rb'))
            #Get number of pages for whole directory
            pdf_pages += pdf.getNumPages()
            #Pages for each pdf file
            df2 = pd.DataFrame([[f, os.path.join(root, f), pdf.getNumPages()]], columns=['fileName', 'fileLocation', 'pageNumber'])
            #Append report with each file
            df = df.append(df2, ignore_index=True)
        

AllPages = "Sum of all pages |" + str(pdf_pages) +"\n\nBelow list of all documents\n \n"
#Save report
Report.writelines(AllPages)
Report.writelines(df.to_csv(columns=['fileName', 'fileLocation', 'pageNumber'], sep='|',index_label="Index", line_terminator="\n"))
Report.close()
#Ending code message
print('Policzono wszystkie strony plików pdf pod ścieżką: ' + dirPathTo + '.\nGdzie również zapisno raport końcowy o nazwie: Count_pages_report.txt')
