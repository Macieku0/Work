import pandas as pd
import os
from PyPDF2 import PdfFileReader
pdf_pages = 0
df = pd.DataFrame(columns=['fileName', 'fileLocation', 'pageNumber'])
for root, dirs, files in os.walk("C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/PDF/"):
    for f in files:
        pdf = PdfFileReader(open(os.path.join(root, f), 'rb'))
        pdf_pages += pdf.getNumPages()
        df2 = pd.DataFrame([[f, os.path.join(root, f), pdf.getNumPages()]], columns=['fileName', 'fileLocation', 'pageNumber'])
        df = df.append(df2, ignore_index=True)
print(df.head)
print(pdf_pages)