import pandas as pd
import numpy as np
import math
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from functools import reduce
from Combinations import CombinationsList
import os
from datetime import date

#Tworzenie folderu w podanej ścieżce i nadpisanie ścieżki generowania raportu
def create_folder(dir):
    dir_path = f'{dir}{date.today()}_report'
    x = 0 
    while True:
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
            dir_path = f'{dir_path}\\'
            break
        elif not os.path.exists(f'{dir_path}_{x}'):
            os.mkdir(f'{dir_path}_{x}')
            dir_path = f'{dir_path}_{x}\\'
            break
        else:
            x += 1
    return dir_path


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
    if x[0] == "/":
        return x[1:len(x)]
    else:
        return x


#Obliczanie sił lokalnych na podstawie wartości globalnych i ułożenia w płaszczyźnie
#Oznaczenia wartośći w tablicy [x] jak poniżej
#'FX' - 0,'FY' - 1,'ORIANGLE' - 2,'COS' - 3,'SIN' - 4
def force(x): 
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


def localDisplacements(x): 
    if ((float(x[2]) + 360)%360) >= 180:
        DxCos = (float(x[0]) * float(x[3]))
        DxSin = (float(x[0]) * float(x[4]))
        DyCos = (float(x[1]) * float(-x[3]))
        DySin = (float(x[1]) * float(x[4]))
    else:
        DxCos = (float(x[0]) * float(x[3]))
        DxSin = (float(x[0]) * float(-x[4]))
        DyCos = (float(x[1]) * float(x[3]))
        DySin = (float(x[1]) * float(x[4]))
    x['DA'] = float('{:.2f}'.format(DxCos + DySin))
    x['DL'] = float('{:.2f}'.format(DxSin + DyCos))
    return x[['DA','DL']]

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
    InitialDir = root.filename


#Wybieranie ścieżki dla raportu AutoPipe
def GetAutoPipe():
    global InitialDir  
    root.filename = filedialog.askopenfilename(initialdir=InitialDir,title='Choose your AutoPipe report file', filetypes=[('XLS','*.xlsx')])
    if root.filename != '':
        AutoPipeEntry.delete(0,END)
        AutoPipeEntry.insert(0,root.filename)
    InitialDir = root.filename


#Wybieranie ścieżki dla folderu
def GetFolderDir():
    global InitialDir  
    root.filename = filedialog.askdirectory(initialdir=InitialDir,title='Choose folder Directory',)
    if root.filename != '':
        PathEntry.delete(0,END)
        PathEntry.insert(0,root.filename + '/')
    InitialDir = root.filename


#Funkcja wyciągająca nazwy kombinacji które zostały wybrane przez użytkownika i podaje w osobnym oknie
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


# Funkcja genrująca kombinacje z pliku autopipe jako forma wyboru dla użytkowników
def OpenComWindow():
    global AllCom
    AutoPipe = AutoPipeEntry.get()
    #Sprawdzanie czy została podana ścieżka do pliku autopipe
    if AutoPipe == '':
        messagebox.showerror('Brak danych!','Podaj ścieżkę do raportu z AutoPipe')
        return
    
    #Generuownie listy kombinacji z pliku
    list = CombinationsList(AutoPipe)

    #Tworzenie okienka z kombinacjami
    ComWindow = Toplevel(root)
    ComWindow.title('Choose Combinations')
    ComWindow.geometry('400x250')

    #Tworzenie checkboxów w trzech kolumnach
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
    Button(ComWindow,text='Confirm choice',command=ChooseCom).grid(column=0,columnspan=3)


#--------------------------------------------------------------------------------------------------------------------------
#Główny program liczący

if __name__ == '__main__':
    def startProgram():

        #Sprawdzenie czy użytkonik wygbrał kombinacje do dalszej pracy programu
        if ChoosedCom == None:
            messagebox.showerror('Brak danych!','Kombinacje nie zostały wybrane')
            return

        #Pobranie wyrbanych przez użytkownika ścieżek do plików i do zapisu raportu
        Path = PathEntry.get()
        PDMS = PDMSEntry.get()
        AutoPipe = AutoPipeEntry.get()


        #Stworzenie folderu i nadpisanie ścieżki
        Path = create_folder(Path)

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
        
        #Lista do sprawdzenia braków
        pdmsSup = PdmsFile['NAME'].unique()


        #Obliczanie sinusów i cosinusów
        PdmsFile['SIN'] = [sin(oriangle) for oriangle in PdmsFile['ORIANGLE']]
        PdmsFile['COS'] = [cos(oriangle) for oriangle in PdmsFile['ORIANGLE']]


        #Określanie czy rurociąg jest poziomy czy pionowy
        PdmsFile = PdmsFile[PdmsFile['ADIR'].notnull()]
        PdmsFile = PdmsFile[PdmsFile['ADIR'].notna()]
        PdmsFile['Vertical'] = ["YES" if (x[:1] in ["D","U"]) else "NO" for x in PdmsFile['ADIR'] ]


        #Zmiana nazwy kolumny
        PdmsFile.rename(columns={'DTXR':'Description','Position WRT Owner':'WRT','NAME':'Name'},inplace=True)

        #Obrabianie Koordynatów
        PdmsFile[['PDMS - CoordX','PDMS - CoordY','PDMS - CoordZ']] = [[x[:x.find('mm')],x[x.find('mm')+3:x.find('mm',x.find('mm')+3)],x[x.find('mm',x.find('mm')+3)+3:len(x)-2]] 
        for x in PdmsFile['WRT']]
        PdmsFile['PDMS - CoordX'] = [x[x.find(' '):] for x in PdmsFile['PDMS - CoordX']]
        PdmsFile['PDMS - CoordY'] = [x[x.find(' '):] for x in PdmsFile['PDMS - CoordY']]
        PdmsFile['PDMS - CoordZ'] = [x[x.find(' '):] for x in PdmsFile['PDMS - CoordZ']]
        PdmsFile[['PDMS - CoordX','PDMS - CoordY','PDMS - CoordZ']] = PdmsFile[['PDMS - CoordX','PDMS - CoordY','PDMS - CoordZ']].astype(float).astype(int)


        #Wczytanie pliku autopipe
        AutoPipeFile = pd.read_excel(f'{AutoPipe}')


        #Selekcja kolumn
        AutoPipeFile = AutoPipeFile[['Tag No.','Type','Combination','GlobalFX','GlobalFY','GlobalFZ','CoordX','CoordY','CoordZ','GlobalDX','GlobalDY','GlobalDZ']]

        #Usuwanie pierwszego wiersza - wiersza z jednostkami
        AutoPipeFile = AutoPipeFile.drop(0)

        #Filtrowanie po wybranych przez użytkownika kombinacjach obliczeniowych
        AutoPipeFile = AutoPipeFile[AutoPipeFile['Combination'].isin(Com)]

        #Zmiana danych na liczbowe, zmiennoprzecinkowe
        AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']] = AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']].astype(float)

        #Zmiana nazw kolummn
        AutoPipeFile.rename(columns={'Tag No.':'Name',
        'GlobalFZ':'FZ',
        'GlobalFX':'FX',
        'GlobalFY':'FY',
        'CoordX':'AutoPipe - CoordX',
        'CoordY':'AutoPipe - CoordY',
        'CoordZ':'AutoPipe - CoordZ'},inplace=True)


        #Czyszczenie nazw z "\"
        AutoPipeFile['Name'] = AutoPipeFile['Name'].astype(str)
        AutoPipeFile['Name'] = [clean_name(name) for name in AutoPipeFile['Name']]
        #Selekcja kolumn do matrycy
        AutoPipeFileCoord = AutoPipeFile[['Name','AutoPipe - CoordX','AutoPipe - CoordY','AutoPipe - CoordZ']]

        #Do dalszego przeliczania przesunięć
        AutoPipeFileDisplecements = AutoPipeFile[['Name','Type','GlobalDX','GlobalDY','GlobalDZ','Combination']]
        AutoPipeFileDisplecements[['GlobalDX','GlobalDY','GlobalDZ']] = AutoPipeFileDisplecements[['GlobalDX','GlobalDY','GlobalDZ']].astype(float)
        print(AutoPipeFileDisplecements.info())
        AutoPipeFileDisplecementsType = AutoPipeFileDisplecements[['Name','Type']].groupby('Name').sum().rename(columns={'Type':'AllType'}).reset_index()

        autoPipeSup = AutoPipeFile['Name'].unique()

        brakPdms = [x for x in autoPipeSup if x not in pdmsSup]
        brakAutoPipe = [x for x in pdmsSup if x not in autoPipeSup]
        print('Brak w PDMS')
        print(brakPdms)
        print('Brak w AutoPipe')
        print(brakAutoPipe)

        #Sumowanie wartości dla kombinacji obliczeniowej dla każdego zamocowania
        AutoPipeFile = AutoPipeFile.groupby(['Name','Combination']).sum().reset_index()


        #Wycinanie wartości koordynatów tak aby zostały tylko wartości liczbowe
        AutoPipeFileCoord['AutoPipe - CoordX'] = [np.absolute(int(x[:str(x).find('.')])) for x in AutoPipeFileCoord['AutoPipe - CoordX']]
        AutoPipeFileCoord['AutoPipe - CoordY'] = [np.absolute(int(x[:str(x).find('.')])) for x in AutoPipeFileCoord['AutoPipe - CoordY']]
        AutoPipeFileCoord['AutoPipe - CoordZ'] = [np.absolute(int(x[:str(x).find('.')])) for x in AutoPipeFileCoord['AutoPipe - CoordZ']]
        AutoPipeFileCoord[['AutoPipe - CoordX','AutoPipe - CoordY','AutoPipe - CoordZ']].astype(int)


        #MAX,MIN,EXTREMUM Colums
        #Na podstawie wybranych opcji (min,max,ext) wybór dla każdego zamocowania odpowiedniej wartości
        Conditions = [Extremum.get(),Maximum.get(),Minimum.get()]
        allDf = []
        if Extremum.get() != '':
            #Wybór ekstremalnej wartości dla każdej osi
            #Zwracanie wartości bezwzględej
            AutoPipeFileEXT = AutoPipeFile.copy()
            AutoPipeFileEXT['ABS(FX)'] = [np.absolute(x) for x in AutoPipeFile['FX']]
            AutoPipeFileEXT['ABS(FY)'] = [np.absolute(x) for x in AutoPipeFile['FY']]
            AutoPipeFileEXT['ABS(FZ)'] = [np.absolute(x) for x in AutoPipeFile['FZ']]
            AutoPipeFileFxEXT = AutoPipeFileEXT[['Name','ABS(FX)']].groupby('Name').max().reset_index()
            AutoPipeFileFxEXT = pd.merge(AutoPipeFileFxEXT,AutoPipeFileEXT[['FX','Combination','ABS(FX)']],on='ABS(FX)',how='left').rename(columns={'Combination':'Extremum - CombinationFx','FX':'Extremum - FX'}).drop_duplicates('Name')
            AutoPipeFileFyEXT = AutoPipeFileEXT[['Name','ABS(FY)']].groupby('Name').max().reset_index()
            AutoPipeFileFyEXT = pd.merge(AutoPipeFileFyEXT,AutoPipeFileEXT[['FY','Combination','ABS(FY)']],on='ABS(FY)',how='left').rename(columns={'Combination':'Extremum - CombinationFy','FY':'Extremum - FY'}).drop_duplicates('Name')
            AutoPipeFileFzEXT = AutoPipeFileEXT[['Name','ABS(FZ)']].groupby('Name').max().reset_index()
            AutoPipeFileFzEXT = pd.merge(AutoPipeFileFzEXT,AutoPipeFileEXT[['FZ','Combination','ABS(FZ)']],on='ABS(FZ)',how='left').rename(columns={'Combination':'Extremum - CombinationFz','FZ':'Extremum - FZ'}).drop_duplicates('Name')
            allDf.extend([AutoPipeFileFxEXT,AutoPipeFileFyEXT,AutoPipeFileFzEXT])
        if Minimum.get() != '':
            AutoPipeFileFxMIN = AutoPipeFile[['Name','FX']].groupby('Name').min().reset_index()
            AutoPipeFileFxMIN = pd.merge(AutoPipeFileFxMIN,AutoPipeFile[['FX','Combination']],on='FX',how='left').rename(columns={'Combination':'Minimum - CombinationFx','FX':'Minimum - FX'}).drop_duplicates('Name')
            AutoPipeFileFyMIN = AutoPipeFile[['Name','FY']].groupby('Name').min().reset_index()
            AutoPipeFileFyMIN = pd.merge(AutoPipeFileFyMIN,AutoPipeFile[['FY','Combination']],on='FY',how='left').rename(columns={'Combination':'Minimum - CombinationFy','FY':'Minimum - FY'}).drop_duplicates('Name')
            AutoPipeFileFzMIN = AutoPipeFile[['Name','FZ']].groupby('Name').min().reset_index()
            AutoPipeFileFzMIN = pd.merge(AutoPipeFileFzMIN,AutoPipeFile[['FZ','Combination']],on='FZ',how='left').rename(columns={'Combination':'Minimum - CombinationFz','FZ':'Minimum - FZ'}).drop_duplicates('Name')
            allDf.extend([AutoPipeFileFxMIN,AutoPipeFileFyMIN,AutoPipeFileFzMIN])
        if Maximum.get() != '':
            AutoPipeFileFxMAX = AutoPipeFile[['Name','FX']].groupby('Name').max().reset_index()
            AutoPipeFileFxMAX = pd.merge(AutoPipeFileFxMAX,AutoPipeFile[['FX','Combination']],on='FX',how='left').rename(columns={'Combination':'Maximum - CombinationFx','FX':'Maximum - FX'}).drop_duplicates('Name')
            AutoPipeFileFyMAX = AutoPipeFile[['Name','FY']].groupby('Name').max().reset_index()
            AutoPipeFileFyMAX = pd.merge(AutoPipeFileFyMAX,AutoPipeFile[['FY','Combination']],on='FY',how='left').rename(columns={'Combination':'Maximum - CombinationFy','FY':'Maximum - FY'}).drop_duplicates('Name')
            AutoPipeFileFzMAX = AutoPipeFile[['Name','FZ']].groupby('Name').max().reset_index()
            AutoPipeFileFzMAX = pd.merge(AutoPipeFileFzMAX,AutoPipeFile[['FZ','Combination']],on='FZ',how='left').rename(columns={'Combination':'Maximum - CombinationFz','FZ':'Maximum - FZ'}).drop_duplicates('Name')
            allDf.extend([AutoPipeFileFxMAX,AutoPipeFileFyMAX,AutoPipeFileFzMAX])


        #Łączenie sił ze wszystkich płaszczyzn do jednej tablicy
        MergedData = reduce(lambda x,y: pd.merge(x,y,on=['Name'],how='outer'),allDf)

        #Usuwanie duplikatów
        MergedData.drop_duplicates('Name')
 

        #Łączenie danych z PDMS'a i AutoPipe'a
        FinalList = ['Name','Description','Vertical']
        for x in Conditions:
            if x != '':
                FinalList.extend([f'{x} - CombinationFx',f'{x} - FX',f'{x} - CombinationFy',f'{x} - FY',f'{x} - CombinationFz',f'{x} - FZ'])
        FinalReport = pd.merge(MergedData,PdmsFile[['Name','ORIANGLE','SIN','COS','Description','PDMS - CoordX','PDMS - CoordY','PDMS - CoordZ','Vertical']],on='Name',how='left')


        #Przeliczanie wartości lokalnych dla danych globalnych i kąta nachylenia w płaszczyźnie X i Y
        #MAX,MIN,EXTREMUM Colums
        allDf = []
        allDf.append(FinalReport)
        if Extremum.get() != '':
            FinalReportEXT = FinalReport.copy()
            FinalReportEXT[['Extremum - FA','Extremum - FL']] = FinalReport[['Extremum - FX','Extremum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportEXT['Extremum - FV'] = FinalReport['Extremum - FZ']
            FinalReportEXT = FinalReportEXT[['Name','Extremum - FA','Extremum - FL','Extremum - FV']]
            allDf.append(FinalReportEXT)
        if Minimum.get() != '':
            FinalReportMIN = FinalReport.copy()
            FinalReportMIN[['Minimum - FA','Minimum - FL']] = FinalReport[['Minimum - FX','Minimum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportMIN['Minimum - FV'] = FinalReport['Minimum - FZ']
            FinalReportMIN = FinalReportMIN[['Name','Minimum - FA','Minimum - FL','Minimum - FV']]
            allDf.append(FinalReportMIN)
        if Maximum.get() != '':
            FinalReportMAX = FinalReport.copy()
            FinalReportMAX[['Maximum - FA','Maximum - FL']] = FinalReport[['Maximum - FX','Maximum - FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
            FinalReportMAX['Maximum - FV'] = FinalReport['Maximum - FZ']
            FinalReportMAX = FinalReportMAX[['Name','Maximum - FA','Maximum - FL','Maximum - FV']]
            allDf.append(FinalReportMAX)
        allDf.append(AutoPipeFileCoord)

        verticalPipes = []
        #Łączenie raportów z siłami lokalnymi do głównej tablicy
        FinalReport = reduce(lambda x,y: pd.merge(x,y,on='Name',how='left'),allDf).drop_duplicates('Name')
        for x in Conditions:
            if x != '':
                FinalList.extend([f'{x} - FL',f'{x} - FA',f'{x} - FV'])
                verticalPipes.extend([f'{x} - FL',f'{x} - FA',f'{x} - FV'])

        #Tworzenie kolumn z deltą koordynatów PDMS vs AutoPipe
        FinalReport['Difference  - CoordX'] = np.absolute(FinalReport['AutoPipe - CoordX'] - FinalReport['PDMS - CoordX'])
        FinalReport['Difference  - CoordY'] = np.absolute(FinalReport['AutoPipe - CoordY'] - FinalReport['PDMS - CoordY'])
        FinalReport['Difference  - CoordZ'] = np.absolute(FinalReport['AutoPipe - CoordZ'] - FinalReport['PDMS - CoordZ'])
        FinalList.extend(['PDMS - CoordX',
        'PDMS - CoordY',
        'PDMS - CoordZ',
        'AutoPipe - CoordX',
        'AutoPipe - CoordY',
        'AutoPipe - CoordZ',
        'Difference  - CoordX',
        'Difference  - CoordY',
        'Difference  - CoordZ'])

        FinalReport.loc[FinalReport.Vertical == 'YES',verticalPipes] = np.nan

        #Przesunięcia

        AutoPipeFileMinDX = AutoPipeFileDisplecements[['Name','GlobalDX']].groupby('Name').min().reset_index()
        AutoPipeFileMinDX = pd.merge(AutoPipeFileMinDX,AutoPipeFileDisplecements[['GlobalDX','Combination']],on='GlobalDX',how='left').rename(columns={'Combination':'Min - CombDX','GlobalDX':'Min - DX'}).drop_duplicates('Name')
        AutoPipeFileMinDY = AutoPipeFileDisplecements[['Name','GlobalDY']].groupby('Name').min().reset_index()
        AutoPipeFileMinDY = pd.merge(AutoPipeFileMinDY,AutoPipeFileDisplecements[['GlobalDY','Combination']],on='GlobalDY',how='left').rename(columns={'Combination':'Min - CombDY','GlobalDY':'Min - DY'}).drop_duplicates('Name')
        AutoPipeFileMinDZ = AutoPipeFileDisplecements[['Name','GlobalDZ']].groupby('Name').min().reset_index()
        AutoPipeFileMinDZ = pd.merge(AutoPipeFileMinDZ,AutoPipeFileDisplecements[['GlobalDZ','Combination']],on='GlobalDZ',how='left').rename(columns={'Combination':'Min - CombDZ','GlobalDZ':'Min - DZ'}).drop_duplicates('Name')
        allDf = []
        allDf.extend([AutoPipeFileMinDX,AutoPipeFileMinDY,AutoPipeFileMinDZ])
        
        AutoPipeFileMaxDX = AutoPipeFileDisplecements[['Name','GlobalDX']].groupby('Name').max().reset_index()
        AutoPipeFileMaxDX = pd.merge(AutoPipeFileMaxDX,AutoPipeFileDisplecements[['GlobalDX','Combination']],on='GlobalDX',how='left').rename(columns={'Combination':'Max - CombDX','GlobalDX':'Max - DX'}).drop_duplicates('Name')
        AutoPipeFileMaxDY = AutoPipeFileDisplecements[['Name','GlobalDY']].groupby('Name').max().reset_index()
        AutoPipeFileMaxDY = pd.merge(AutoPipeFileMaxDY,AutoPipeFileDisplecements[['GlobalDY','Combination']],on='GlobalDY',how='left').rename(columns={'Combination':'Max - CombDY','GlobalDY':'Max - DY'}).drop_duplicates('Name')
        AutoPipeFileMaxDZ = AutoPipeFileDisplecements[['Name','GlobalDZ']].groupby('Name').max().reset_index()
        AutoPipeFileMaxDZ = pd.merge(AutoPipeFileMaxDZ,AutoPipeFileDisplecements[['GlobalDZ','Combination']],on='GlobalDZ',how='left').rename(columns={'Combination':'Max - CombDZ','GlobalDZ':'Max - DZ'}).drop_duplicates('Name')

        allDf.extend([AutoPipeFileMaxDX,AutoPipeFileMaxDY,AutoPipeFileMaxDZ])

        MergedDisplacements = reduce(lambda x,y: pd.merge(x,y,on=['Name'],how='outer'),allDf).drop_duplicates('Name')

        FinalReport = pd.merge(FinalReport,MergedDisplacements[['Name','Min - DX','Min - DY','Min - DZ','Max - DX','Max - DY','Max - DZ']],on='Name',how='left')

        FinalReport = pd.merge(FinalReport,AutoPipeFileDisplecementsType[['Name','AllType']], on='Name',how='left')

        allDf = []

        FinalReportMIN = FinalReport.copy()
        FinalReportMIN[['Min - DA','Min - DL']] = FinalReportMIN[['Min - DX','Min - DY','ORIANGLE','COS','SIN',]].apply(localDisplacements,axis=1)
        FinalReportMIN['Min - DV'] = FinalReportMIN['Min - DZ']
        FinalReportMIN = FinalReportMIN[['Name','Min - DA','Min - DL','Min - DV']]

        FinalReportMAX = FinalReport.copy()
        FinalReportMAX[['Max - DA','Max - DL']] = FinalReportMAX[['Max - DX','Max - DY','ORIANGLE','COS','SIN',]].apply(localDisplacements,axis=1)
        FinalReportMAX['Max - DV'] = FinalReportMAX['Max - DZ']
        FinalReportMAX = FinalReportMAX[['Name','Max - DA','Max - DL','Max - DV']]


        allDf.extend([FinalReport,FinalReportMIN,FinalReportMAX])

        FinalReport = reduce(lambda x,y: pd.merge(x,y,on=['Name'],how='left'),allDf).drop_duplicates('Name')
        FinalList.extend(['Min - DA','Min - DL','Min - DV','Max - DA','Max - DL','Max - DV'])

        FinalReport.loc[FinalReport['AllType'].str.contains('Guide'),'Max - DL'] = 0
        FinalReport.loc[FinalReport['AllType'].str.contains('Guide'),'Min - DL'] = 0
        FinalReport.loc[FinalReport['AllType'].str.contains('Line Stp'),'Min - DA'] = 0
        FinalReport.loc[FinalReport['AllType'].str.contains('Line Stp'),'Max - DA'] = 0

        FinalReport.to_excel(f'{Path}Caly_{Out}')


        #Tworzenie formy raportu końcowego
        FinalReport = FinalReport[FinalList]

        #Usunięcie duplikatów
        FinalReport = FinalReport.drop_duplicates('Name').reset_index(drop=True)

        # Wygenerowanie raportu końcowego
        FinalReport.to_excel(f'{Path}{Out}')

        #Wiadomość końcowa po wygenerowaniu raportu
        messagebox.showinfo('Raport Gotowy!',f'Plik został zapisany pod scieżką: {Path}{Out}')
#Koniec działania główniej funkcji programu
#--------------------------------------------------------------------------------------------------------------------------
#Okienko programu


    #Tkinter start - definiowanie tytułu i wymiarów okienka
    root = Tk()
    root.title('Report Creator')
    root.geometry('550x175')
   

    #Okienko do wprowadzania ścieżki do raportu z PDMS'a
    PDMSEntry = Entry(root, width=55,borderwidth=2)
    PDMSEntry.grid(column=1,columnspan=2,row=1)
    PdmsButton = Button(root, text='Get PDMS report directory',command=GetPDMS,width=25).grid(column=0,row=1)


    #Okienko do wprowadzania ścieżki do raportu z AutoPipe'a
    AutoPipeEntry = Entry(root, width=55,borderwidth=2)
    AutoPipeEntry.grid(column=1,columnspan=2,row=2)
    AutoPipeButton = Button(root, text='Get AutoPipe report directory',command=GetAutoPipe,width=25).grid(column=0,row=2)


    #Okienko do wprowadzania ścieżki gdzie zostanie wygenerowany raport
    PathEntry = Entry(root, width=55,borderwidth=2)
    PathEntry.grid(column=1,columnspan=2,row=0)
    PathButton = Button(root, text='Get folder directory',command=GetFolderDir,width=25).grid(column=0,row=0)


    #Dla wybranej opcji program przelicza Fx Fy Fz tak jak teraz tylko dodaje kolumny dodatkowo dla max i min - 3 odzielne check-boxy 
    global Maximum
    global Minimum
    global Extremum
    Minimum = StringVar()
    Maximum = StringVar()
    Extremum = StringVar()
    Checkbutton(root,text='MIN',variable=Minimum, onvalue='Minimum',offvalue='').grid(column=0,row=3)
    Checkbutton(root,text='MAX',variable=Maximum, onvalue='Maximum',offvalue='').grid(column=1,row=3)
    Checkbutton(root,text='EXTREME',variable=Extremum, onvalue='Extremum',offvalue='').grid(column=2,row=3)

    #Przycisk generujący kombinacje do wybrania
    ComWindowButton = Button(root,text='Choose combinations',command=OpenComWindow).grid(column=0,columnspan=3,row=4)

    #Przycisk startu programu
    StartButton = Button(root,text='Start',command=startProgram,height=1,width=10).grid(column=0,columnspan=3)
    root.mainloop()