import openpyxl
import pandas
import os

#Main files directory
MainDir = "C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/20210519_LISTA_WPALEN_python/"
#List name
ListName = "LISTA_WPALEK.xlsx"
#Name of list worksheet
WorkSheetName = "Tie-In list"
#Letter of column with names for naming files
Col_lett = "K"
#Range of column with names of files
ColStart = 12
ColEnd = 24
#Template name
TemplateName = "Karta technologiczna_wpalki.xlsx"
#Name of template worksheet 
TemplateWorkSheetName = "KARTA"
#Files save / update directory
FilesDir = "C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/20210519_LISTA_WPALEN_python/WYDRUKI/"
#List of files to create / update
FilesList = []
#Map of cells connection between list and template
Map = {'A7':'C','B7':'G','C7':'D'}

#Creating a list of all files name
for x in range(ColStart,ColEnd):
    wb = openpyxl.load_workbook(MainDir + ListName)
    worksheet = wb[WorkSheetName]
    FilesList.append(worksheet[Col_lett + str(x)].value)
    wb.close()

#Creating files for each element in the list
source = openpyxl.load_workbook(MainDir + TemplateName)
for file in FilesList:
    #Checking if file exist
    if os.path.isfile(FilesDir + file + ".xlsx"):
        print(file + ".xlsx" + " already exist")
    else:
        #If no so create one
        source.save(FilesDir + file + ".xlsx")
        source.close()

#Fullfil created files with data from list
i = 0
for z in range(ColStart,ColEnd):
    wb = openpyxl.load_workbook(MainDir + ListName)
    wbWroksheet = wb[WorkSheetName]
    NewFile = openpyxl.load_workbook(FilesDir + FilesList[i] + ".xlsx")
    NewWorksheet = NewFile[TemplateWorkSheetName]
    for x,y in Map.items():
        NewWorksheet[x].value = wbWroksheet[y + str(z)].value
        NewFile.save(FilesDir + FilesList[i] + ".xlsx")
        NewFile.close()
    i += 1
    wb.close()