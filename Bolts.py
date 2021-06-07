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
        indexList = ['','','','','']
        i = 0
        for line in lines:
            if not line.isspace():
                a = str(line)
                if (len(a) <= 111 and len(a) >= 106):
                    description = a[2:40]
                    itemCode = a[70:92]
                    quantity = re.sub('\n','',a[108:111])
                if (len(a) <= 50 and len(a) >= 10):
                    if a[1:9] == 'PIPELINE':
                        indexList = ['','','','','']
                        pipeline = re.sub('\n','',a[14:len(a)])
                        indexList[0] = pipeline
                    elif a[5:7] == a.split()[0][0:2]:
                        secondDesc = re.sub('\n','',a[5:len(a)])
                        description = description + ', ' + secondDesc
                        indexList[1] = description
                        indexList[2] = itemCode
                        indexList[3] = quantity
                        indexList[4] = secondDesc
                        i += 1
                        finalList.append(indexList.copy())
                if (len(a) <= 10 and len(a) >= 0):
                    material = re.sub('\n','',a[5:len(a)])
                    indexList[1] = description + ', ' + material
                    indexList[4] = material
                    del finalList[i-1]
                    finalList.append(indexList.copy())

    src = pd.DataFrame(finalList,columns=['PIPLINE NAME','DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL'])
    src['SECTION'] = 'ŚRUBY, NAKRĘTKI'
    src['DN1'] = [x[:3] for x in src['ITEM-CODE']]
    src['QUANTITY'] = [int(x) for x in src['QUANTITY']]

    poRurach = src.copy()
    poRurach = poRurach[['PIPLINE NAME','ITEM-CODE','DESCRIPTION','QUANTITY','MATERIAL']].groupby(['PIPLINE NAME','ITEM-CODE','DESCRIPTION','MATERIAL']).sum().reset_index()


    zbiorowe = src.copy()
    zbiorowe = zbiorowe[['DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL']].groupby(['ITEM-CODE','DESCRIPTION','MATERIAL']).sum().reset_index()


    zbiorowe = pd.merge(zbiorowe,src[['ITEM-CODE','SECTION','DN1']],on=['ITEM-CODE'],how='outer').drop_duplicates(['ITEM-CODE','DESCRIPTION']).reset_index()
    poRurach = pd.merge(poRurach,src[['ITEM-CODE','SECTION','DN1']],on=['ITEM-CODE'],how='outer').drop_duplicates(['PIPLINE NAME','ITEM-CODE','DESCRIPTION']).reset_index()

    #Creating xlsx file
    writer = pd.ExcelWriter(pathTo)

    #Save to xlsx file
    poRurach[['SECTION','PIPLINE NAME','ITEM-CODE','DN1','DESCRIPTION','MATERIAL','QUANTITY']].to_excel(writer, sheet_name='Po rurociągach')
    zbiorowe[['SECTION','ITEM-CODE','DN1','DESCRIPTION','MATERIAL','QUANTITY']].to_excel(writer, sheet_name='Zbiorowe')

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