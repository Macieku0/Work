from numpy.core.fromnumeric import sort
import pandas as pd
import numpy as np
import math
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PIL import ImageTk, Image
from timeit import default_timer
import pathlib
from functools import reduce
from Combinations import CombinationsList

#Obliczanie cosinusa z podanego kąta
def cos(x):
    return math.cos(math.radians(float(x))) 
#Obliczanie sinusa z podanego kąta
def sin(x):
    return math.sin(math.radians(float(x)))
#Czyszczenie katów z dopiska degree
def clean_angles(x):
    return x[x.rfind(' ')+1:len(x)]
#Czyszczenie nazw zamocowań z '/'
def clean_name(x):
    return x[1:len(x)]
#Obliczanie sił lokalnych na podstawie wartości globalnych i ułożenia w płaszczyźnie
def force(x): #'FX' - 0,'FY' - 1,'ORIANGLE' - 2,'COS' - 3,'SIN' - 4
    if ((float(x[2]) + 360)%360) >= 180:
        FxCos = (float(x[0]) * float(x[3]))
        FxSin = (float(x[0]) * float(x[4]))
        FyCos = (float(x[1]) * float(-x[3]))
        FySin = (float(x[1]) * float(x[4]))
    else:
        FxCos = (float(x[0]) * float(x[3]))
        FxSin = (float(x[0]) * float(-x[4]))
        FyCos = (float(x[1]) * float(x[3]))
        FySin = (float(x[1]) * float(x[4]))
    x['FA'] = float('{:.2f}'.format(FxCos + FySin))
    x['FL'] = float('{:.2f}'.format(FxSin + FyCos))
    return x[['FA','FL']]
#Wstępna lokalizacja do wyboru folderów / plików
global InitialDir  
InitialDir = 'C:/'
# Wybieranie ścieżki dla raportu PDMS
def GetPDMS():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose your PDMS report file', filetypes=[('CSV','*.csv')])
    if root.filename != '':
        PDMSEntry.delete(0,END)
        PDMSEntry.insert(0,root.filename)
    #global InitialDir
    InitialDir = root.filename
#Wybieranie ścieżki dla raportu AutoPipe
def GetAutoPipe():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose your AutoPipe report file', filetypes=[('XLS','*.xlsx')])
    if root.filename != '':
        AutoPipeEntry.delete(0,END)
        AutoPipeEntry.insert(0,root.filename)
    #global InitialDir 
    InitialDir = root.filename
#Wybieranie ścieżki dla folderu
def GetFolderDir():
    global InitialDir  
    root.filename = filedialog.askdirectory(initialdir=InitialDir,title='Choose folder Directory',)
    if root.filename != '':
        PathEntry.delete(0,END)
        PathEntry.insert(0,root.filename + '/')
    #global InitialDir 
    InitialDir = root.filename

def ChooseCom():
    global ChoosedCom
    ChoosedCom = []
    for com in AllCom:
        nameOfCom = com.get()
        if nameOfCom != '':
            ChoosedCom.append(nameOfCom)
    ChoosedComWindow = Toplevel(root)
    ChoosedComWindow.title('Wybrane kombinacje')
    ChoosedComWindow.geometry('200x400')
    for x in ChoosedCom:
        x = Label(ChoosedComWindow,text=x).pack()

def OpenComWindow():
    global AllCom
    AutoPipe = AutoPipeEntry.get()
    if AutoPipe == '':
        messagebox.showerror('Brak danych!','Podaj ścieżkę do raportu z AutoPipe')
        return
    list = CombinationsList(AutoPipe)
    ComWindow = Toplevel(root)
    ComWindow.title('Choose Combinations')
    ComWindow.geometry('400x250')
    i = 0
    AllCom =[]
    for x in list:
        strVar = StringVar()
        if i <=6:
            Checkbutton(ComWindow,text=x,variable=strVar, onvalue=x,offvalue='').grid(column=0,row=i)
        elif i <=12:
            Checkbutton(ComWindow,text=x,variable=strVar, onvalue=x,offvalue='').grid(column=1,row=i-7)
        else:
            Checkbutton(ComWindow,text=x,variable=strVar, onvalue=x,offvalue='').grid(column=2,row=i-13)
        i += 1
        AllCom.append(strVar)   
    #Przycisk potwierdzający wybór zaznaczonych kombinacji
    ConfirmCom = Button(ComWindow,text='Confirm choice',command=ChooseCom).grid(column=0,columnspan=3)

if __name__ == '__main__':
    def startProgram():
        if ChoosedCom == None:
            messagebox.showerror('Brak danych!','Kombinacje nie zostały wybrane')
            return
        Path = PathEntry.get()
        PDMS = PDMSEntry.get()
        AutoPipe = AutoPipeEntry.get()
        #Załadowanie kombinacji z drugiego okna
        Com = ChoosedCom
        #Nazwa raportu wyjściowego
        Out = 'Raport_policzony.xlsx'

        #Wczytanie pliku pdms i podział na kolumny
        PdmsFile = pd.read_csv(f'{PDMS}',sep='|',decimal='.')
        #Czyszczenie danych
        PdmsFile = PdmsFile.replace("=","'=",regex=True)
        PdmsFile = PdmsFile.replace('degree','',regex=True)
        PdmsFile['ORIANGLE'] = [clean_angles(oriangle) for oriangle in PdmsFile['ORIANGLE']]
        PdmsFile['NAME'] = [clean_name(name) for name in PdmsFile['NAME']]
        #Obliczanie sinusów i cosinusów
        PdmsFile['SIN'] = [sin(oriangle) for oriangle in PdmsFile['ORIANGLE']]
        PdmsFile['COS'] = [cos(oriangle) for oriangle in PdmsFile['ORIANGLE']]
        #Zmiana nazwy kolumny
        PdmsFile.rename(columns={'DTXR':'Description'},inplace=True)

        #Wczytanie pliku autopipe
        AutoPipeFile = pd.read_excel(f'{AutoPipe}')
        #Selekcja kolumn
        AutoPipeFile = AutoPipeFile[['Tag No.','Combination','GlobalFX','GlobalFY','GlobalFZ']]
        #Usuwanie pierwszego wiersza
        AutoPipeFile = AutoPipeFile.drop(0)
        #Filtrowanie po wybranych kombinacjach obliczeniowych
        AutoPipeFile = AutoPipeFile[AutoPipeFile['Combination'].isin(Com)]
        #Zmiana danych na liczbowe
        AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']] = AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']].astype(float)
        #Zmiana nazw kolummn
        AutoPipeFile.rename(columns={'Tag No.':'NAME','GlobalFZ':'FZ','GlobalFX':'FX','GlobalFY':'FY'},inplace=True)
        #Czyszczenie nazw z "\"
        AutoPipeFile['NAME'] = [clean_name(name) for name in AutoPipeFile['NAME']]
        #Sumowanie wartości dla kombinacji obliczeniowej dla każdego zamocowania
        AutoPipeFile = AutoPipeFile.groupby(['NAME','Combination']).sum().reset_index()


        
        #MAX,MIN,EXTREMUM Colums
        Conditions = [Extremum.get(),Maximum.get(),Minimum.get()]
        allDf = []
        if Extremum.get() != '':
            #Wybór ekstremalnej wartości dla każdej osi
            #Zwracanie wartości bezwzględej
            AutoPipeFileEXT = AutoPipeFile.copy()
            AutoPipeFileEXT['ABS(FX)'] = [np.absolute(x) for x in AutoPipeFile['FX']]
            AutoPipeFileEXT['ABS(FY)'] = [np.absolute(x) for x in AutoPipeFile['FY']]
            AutoPipeFileEXT['ABS(FZ)'] = [np.absolute(x) for x in AutoPipeFile['FZ']]
            AutoPipeFileFxEXT = AutoPipeFileEXT[['NAME','ABS(FX)']].groupby('NAME').max().reset_index()
            AutoPipeFileFxEXT = pd.merge(AutoPipeFileFxEXT,AutoPipeFileEXT[['FX','Combination','ABS(FX)']],on='ABS(FX)',how='left').rename(columns={'Combination':'Extremum - CombinationFx','FX':'Extremum - FX'}).drop_duplicates('NAME')
            AutoPipeFileFyEXT = AutoPipeFileEXT[['NAME','ABS(FY)']].groupby('NAME').max().reset_index()
            AutoPipeFileFyEXT = pd.merge(AutoPipeFileFyEXT,AutoPipeFileEXT[['FY','Combination','ABS(FY)']],on='ABS(FY)',how='left').rename(columns={'Combination':'Extremum - CombinationFy','FY':'Extremum - FY'}).drop_duplicates('NAME')
            AutoPipeFileFzEXT = AutoPipeFileEXT[['NAME','ABS(FZ)']].groupby('NAME').max().reset_index()
            AutoPipeFileFzEXT = pd.merge(AutoPipeFileFzEXT,AutoPipeFileEXT[['FZ','Combination','ABS(FZ)']],on='ABS(FZ)',how='left').rename(columns={'Combination':'Extremum - CombinationFz','FZ':'Extremum - FZ'}).drop_duplicates('NAME')
            allDf.extend([AutoPipeFileFxEXT,AutoPipeFileFyEXT,AutoPipeFileFzEXT])
        if Minimum.get() != '':
            AutoPipeFileFxMIN = AutoPipeFile[['NAME','FX']].groupby('NAME').min().reset_index()
            AutoPipeFileFxMIN = pd.merge(AutoPipeFileFxMIN,AutoPipeFile[['FX','Combination']],on='FX',how='left').rename(columns={'Combination':'Minimum - CombinationFx','FX':'Minimum - FX'}).drop_duplicates('NAME')
            AutoPipeFileFyMIN = AutoPipeFile[['NAME','FY']].groupby('NAME').min().reset_index()
            AutoPipeFileFyMIN = pd.merge(AutoPipeFileFyMIN,AutoPipeFile[['FY','Combination']],on='FY',how='left').rename(columns={'Combination':'Minimum - CombinationFy','FY':'Minimum - FY'}).drop_duplicates('NAME')
            AutoPipeFileFzMIN = AutoPipeFile[['NAME','FZ']].groupby('NAME').min().reset_index()
            AutoPipeFileFzMIN = pd.merge(AutoPipeFileFzMIN,AutoPipeFile[['FZ','Combination']],on='FZ',how='left').rename(columns={'Combination':'Minimum - CombinationFz','FZ':'Minimum - FZ'}).drop_duplicates('NAME')
            allDf.extend([AutoPipeFileFxMIN,AutoPipeFileFyMIN,AutoPipeFileFzMIN])
        if Maximum.get() != '':
            AutoPipeFileFxMAX = AutoPipeFile[['NAME','FX']].groupby('NAME').max().reset_index()
            AutoPipeFileFxMAX = pd.merge(AutoPipeFileFxMAX,AutoPipeFile[['FX','Combination']],on='FX',how='left').rename(columns={'Combination':'Maximum - CombinationFx','FX':'Maximum - FX'}).drop_duplicates('NAME')
            AutoPipeFileFyMAX = AutoPipeFile[['NAME','FY']].groupby('NAME').max().reset_index()
            AutoPipeFileFyMAX = pd.merge(AutoPipeFileFyMAX,AutoPipeFile[['FY','Combination']],on='FY',how='left').rename(columns={'Combination':'Maximum - CombinationFy','FY':'Maximum - FY'}).drop_duplicates('NAME')
            AutoPipeFileFzMAX = AutoPipeFile[['NAME','FZ']].groupby('NAME').max().reset_index()
            AutoPipeFileFzMAX = pd.merge(AutoPipeFileFzMAX,AutoPipeFile[['FZ','Combination']],on='FZ',how='left').rename(columns={'Combination':'Maximum - CombinationFz','FZ':'Maximum - FZ'}).drop_duplicates('NAME')
            allDf.extend([AutoPipeFileFxMAX,AutoPipeFileFyMAX,AutoPipeFileFzMAX])

        #Łączenie wszystkich osi razem do jednego pliku
        MergedData = reduce(lambda x,y: pd.merge(x,y,on=['NAME'],how='outer'),allDf)
        #Usuwanie duplikatów
        MergedData.drop_duplicates('NAME')
        MergedData.to_excel(f'{Path}{Out}') 
 
        #Łączenie danych z PDMS'a i AutoPipe'a
        FinalList = ['NAME','Description']
        for x in Conditions:
            if x != '':
                FinalList.extend([f'{x} - CombinationFx',f'{x} - FX',f'{x} - CombinationFy',f'{x} - FY',f'{x} - CombinationFz',f'{x} - FZ'])

        FinalReport = pd.merge(MergedData,PdmsFile[['NAME','ORIANGLE','SIN','COS','Description']],on='NAME',how='left')
        #Przeliczanie wartości lokalnych dla danych globalnych i kąta nachylenia w płaszczyźnie X Y
        #MAX,MIN,EXTREMUM Colums
        allDf = []
        if Extremum.get() != '':
            FinalReportEXT = FinalReport.copy()
            FinalReportEXT[['Extremum - FA','Extremum - FL']] = FinalReport[['Extremum - FX','Extremum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportEXT['Extremum - FV'] = FinalReport['Extremum - FZ']
            allDf.append(FinalReportEXT)
        if Minimum.get() != '':
            FinalReportMIN = FinalReport.copy()
            FinalReportMIN[['Minimum - FA','Minimum - FL']] = FinalReport[['Minimum - FX','Minimum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportMIN['Minimum - FV'] = FinalReport['Minimum - FZ']
            allDf.append(FinalReportMIN)
        if Maximum.get() != '':
            FinalReportMAX = FinalReport.copy()
            FinalReportMAX[['Maximum - FA','Maximum - FL']] = FinalReport[['Maximum - FX','Maximum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportMAX['Maximum - FV'] = FinalReport['Maximum - FZ']
            allDf.append(FinalReportMAX)
        FinalReportEXT.to_excel('C:\\Dev\\work\\Work_python\\Atta\\test.xlsx')
        FinalReport = reduce(lambda x,y: pd.merge(x,y,on=['NAME'],how='outer'),allDf)
        for x in Conditions:
            if x != '':
                FinalList.extend([f'{x} - FL',f'{x} - FA',f'{x} - FV'])

        FinalReport = FinalReport[FinalList]
        # Wygenerowanie raportu końcowego
        FinalReport.to_excel(f'{Path}{Out}')
        #Wiadomość na koniec generowania raportu
        messagebox.showinfo('Raport Gotowy!',f'Plik został zapisany pod scieżką: {Path}{Out}')



    #Tkinter start

    root = Tk()
    root.title('Report Creator')
    root.geometry('550x175')
   

    #Nazwa raportu PDMS
    PDMSEntry = Entry(root, width=55,borderwidth=2)
    PDMSEntry.grid(column=1,columnspan=2,row=1)
    PdmsButton = Button(root, text='Get PDMS report directory',command=GetPDMS,width=25).grid(column=0,row=1)
    #Nazwa raportu Autopipe
    AutoPipeEntry = Entry(root, width=55,borderwidth=2)
    AutoPipeEntry.grid(column=1,columnspan=2,row=2)
    AutoPipeButton = Button(root, text='Get AutoPipe report directory',command=GetAutoPipe,width=25).grid(column=0,row=2)
    #Scieżka do plików
    PathEntry = Entry(root, width=55,borderwidth=2)
    PathEntry.grid(column=1,columnspan=2,row=0)
    PathButton = Button(root, text='Get folder directory',command=GetFolderDir,width=25).grid(column=0,row=0)
    #Dla wybranej opcji program przelicza Fx Fy Fz tak jak teraz tylko dodaje kolumny dodatkowo dla max i min - 3 odzielne ścieżki
    global Maximum
    global Minimum
    global Extremum
    Minimum = StringVar()
    Maximum = StringVar()
    Extremum = StringVar()
    Checkbutton(root,text='MIN',variable=Minimum, onvalue='Minimum',offvalue='').grid(column=0,row=3)
    Checkbutton(root,text='MAX',variable=Maximum, onvalue='Maximum',offvalue='').grid(column=1,row=3)
    Checkbutton(root,text='EXTREME',variable=Extremum, onvalue='Extremum',offvalue='').grid(column=2,row=3)

    #Kombinacja do wybrania
    ComWindowButton = Button(root,text='Choose combinations',command=OpenComWindow).grid(column=0,columnspan=3,row=4)


    StartButton = Button(root,text='Start',command=startProgram,height=1,width=10).grid(column=0,columnspan=3)
    root.mainloop()



