import pandas as pd
import re
import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog



def main(x):
    pathFrom = 'C:\\PR\\BOLT.TXT'
    pathTo = 'C:\\PR\\BOLT.xlsx'
    with open(pathFrom, 'r', encoding='utf8') as file:
        lines = file.readlines()
        finalList = []
        global indexList
        indexList = ['','','','','']
        i = -1
        for line in lines:
                if not line.isspace():
                    a = str(line)
                    
                    #Nowy rurociąg - dodanie nazwy oraz stworzenie nowej listy
                    if 'PIPELINE' in a:
                            indexList = ['','','','','']
                            pipeline = re.sub('\n','',a[14:len(a)])
                            indexList[0] = pipeline
                            
                    #Jeśli długa linia - pierwszy opis z ilością, opisem i item-codem
                    if (len(a) <= 112 and len(a) >= 106):
                        if a[2] != ' ':
                            #Opis
                            description = re.sub('  ',' ',a[0:44])
                            description = re.sub('  ',' ',description)
                            #Item-code
                            itemCode = re.sub(' ','',a[70:100])
                            #Ilość
                            quantity1 = re.sub(' ','',re.sub('\n','',a[100:108]))
                            quantity2 = re.sub('\n','',a[108:111])
                            #Długość
                            length = re.sub('  ',' ',a[44:57])

                            if quantity1 == '0':
                                quantity = quantity2
                            else:
                                quantity = quantity1
                                
                            indexList[1] = description

                            if itemCode[7:10] == 'AAA' and len(itemCode[21:len(itemCode)]) == 3:
                                itemCode = itemCode[0:7] + itemCode[21:len(itemCode)] + itemCode[10:21]
                            elif itemCode[7:10] == 'AAA' and len(itemCode[21:len(itemCode)]) == 2:
                                itemCode = itemCode[0:7] + itemCode[21:len(itemCode)] + '-' + itemCode[10:21]
                            indexList[2] = itemCode


                            indexList[3] = quantity
                            indexList[4] = length
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


    src = pd.DataFrame(finalList,columns=['PIPLINE NAME','DESCRIPTION','ITEM-CODE','QUANTITY','LENGTH','MATERIAL'])
    src['SECTION'] = 'ŚRUBY, NAKRĘTKI'
    if x == 'yes':
        src['DN1'] = [x[:3] for x in src['ITEM-CODE']]
    else:
        src['DN1'] = src['LENGTH']
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

    root = tk.Tk()
    root.overrideredirect(1)
    root.withdraw()
    messagebox.showinfo('Raport Gotowy!',f'Plik został zapisany pod scieżką: {pathTo}')
    root.destroy()



root = tk.Tk()
root.overrideredirect(1)
root.withdraw()
condition = messagebox.askquestion('Norma','Jeśli ma być wybrana norma EN, kliknij "tak", w innym przypadku klinknij "nie"')
root.destroy()


main(condition)
