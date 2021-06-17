import pandas as pd
import re
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

def main():
    pathFrom = BoltTextFile.get()
    pathTo = f'{PathEntry.get()}BOLT.xlsx'
    with open(pathFrom, 'r', encoding='utf8') as file:
        lines = file.readlines()
        finalList = []
        global indexList
        indexList = ['','','','']
        i = -1
        for line in lines:
                if not line.isspace():
                    a = str(line)
                    
                    #Nowy rurociąg - dodanie nazwy oraz stworzenie nowej listy
                    if 'PIPELINE' in a:
                            indexList = ['','','','']
                            pipeline = re.sub('\n','',a[14:len(a)])
                            indexList[0] = pipeline
                            
                    #Jeśli długa linia - pierwszy opis z ilością, opisem i item-codem
                    if (len(a) <= 112 and len(a) >= 106):
                        #Opis
                        description = re.sub('  ',' ',a[0:44])
                        description = re.sub('  ',' ',description)
                        #Item-code
                        itemCode = re.sub(' ','',a[70:100])
                        #Ilość
                        quantity1 = re.sub(' ','',re.sub('\n','',a[100:108]))
                        quantity2 = re.sub('\n','',a[108:111])
                        
                        if quantity1 == '0':
                            quantity = quantity2
                        else:
                            quantity = quantity1
                            
                        indexList[1] = description
                        indexList[2] = itemCode
                        indexList[3] = quantity
                        
                        finalList.append(indexList.copy())
                        i += 1
                        
                    #Jeśli krótka linia - drugi opis
                    if (len(a) <= 50 and len(a) >= 10):
                        #Sprawdzanie czy to druga linia opisu
                        if a.split()[0][0:2] in a[5:7]:
                            
                            secondDesc = re.sub('  ',' ',re.sub('\n','',a[5:len(a)]))
                            description = f'{description}{secondDesc}'
                            description = re.sub('  ',' ',description)
                            indexList[1] = description
                        
                            del finalList[i]
                            finalList.append(indexList.copy())
                            
                    #Jeśli trzecia linia - trzeci opis
                    if (len(a) <= 10 and len(a) >= 0):
                        
                        thirdDesc = re.sub(' ','',re.sub('\n','',a[5:len(a)]))
                        description = f'{description}{thirdDesc}'
                        
                        description = re.sub('  ',' ',description)
                        indexList[1] = description
                        
                        del finalList[i]
                        finalList.append(indexList.copy())
                    
    for item in finalList:
        description = item[1]
        #Adding material
        item.append(description[description.find(';')+2:])


    src = pd.DataFrame(finalList,columns=['PIPLINE NAME','DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL'])
    src['SECTION'] = 'ŚRUBY, NAKRĘTKI'
    src['DN1'] = [x[:3] for x in src['ITEM-CODE']]
    src['QUANTITY'] = [int(x) for x in src['QUANTITY']]
    src[['PN','THICKNESS','DN2','NAME']] = ['-','-','-','-']
    

    poRurach = src.copy()
    poRurach = poRurach[['PIPLINE NAME','ITEM-CODE','DESCRIPTION','QUANTITY','MATERIAL']].groupby(['PIPLINE NAME','ITEM-CODE','DESCRIPTION','MATERIAL']).sum().reset_index()


    zbiorowe = src.copy()
    zbiorowe = zbiorowe[['DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL']].groupby(['ITEM-CODE','DESCRIPTION','MATERIAL']).sum().reset_index()


    zbiorowe = pd.merge(zbiorowe,src[['PN','THICKNESS','DN2','NAME','ITEM-CODE','SECTION','DN1']],on=['ITEM-CODE'],how='outer').drop_duplicates(['ITEM-CODE','DESCRIPTION']).reset_index()
    poRurach = pd.merge(poRurach,src[['PN','THICKNESS','DN2','NAME','ITEM-CODE','SECTION','DN1']],on=['ITEM-CODE'],how='outer').drop_duplicates(['PIPLINE NAME','ITEM-CODE','DESCRIPTION']).reset_index()

    #Creating xlsx file
    writer = pd.ExcelWriter(pathTo)

    #Save to xlsx file
    poRurach[['PIPLINE NAME','SECTION','ITEM-CODE','PN','MATERIAL','THICKNESS','DN1','DN2','DESCRIPTION','QUANTITY']].to_excel(writer, sheet_name='Po rurociągach')
    zbiorowe[['SECTION','ITEM-CODE','PN','MATERIAL','THICKNESS','DN1','DN2','NAME','DESCRIPTION','QUANTITY']].to_excel(writer, sheet_name='Zbiorowe')

    #Write xlsx file
    writer.save()
    messagebox.showinfo('Raport Gotowy!',f'Plik został zapisany pod scieżką: {pathTo}')


global InitialDir  
InitialDir = 'C:\\PR\\'

#Tkinter start
def GetFile():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose your PDMS report file', filetypes=[('TEXT','*.txt')])
    if root.filename != '':
        BoltTextFile.delete(0,END)
        BoltTextFile.insert(0,root.filename)
    #global InitialDir
    InitialDir = root.filename
def GetFolderDir():
    global InitialDir  
    root.filename = filedialog.askdirectory(initialdir=InitialDir,title='Choose folder Directory',)
    if root.filename != '':
        PathEntry.delete(0,END)
        PathEntry.insert(0,root.filename + '/')
    #global InitialDir 
    InitialDir = root.filename

root = Tk()
root.title('Bolts MTO Creator')
root.geometry('400x75')

#Nazwa raportu PDMS
BoltTextFile = Entry(root, width=40,borderwidth=2)
BoltTextFile.insert(0,'C:\\PR\\BOLT.txt')
BoltTextFile.grid(column=1,row=0)
GetBoltPathButton = Button(root, text='Get path to .txt file',command=GetFile,width=20).grid(column=0,row=0)

#Scieżka do plików
PathEntry = Entry(root, width=40,borderwidth=2)
PathEntry.insert(0,'C:\\PR\\')
PathEntry.grid(column=1,row=1)
PathButton = Button(root, text='Get folder directory',command=GetFolderDir,width=20).grid(column=0,row=1)

StartButton = Button(root,text='Start',command=main,height=1,width=20).grid(column=0,columnspan=2)
root.mainloop()