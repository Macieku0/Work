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

def cos(x):
    return math.cos(math.radians(float(x))) 
def sin(x):
    return math.sin(math.radians(float(x)))
def clean_angles(x):
    return x[x.rfind(' ')+1:len(x)]
def clean_name(x):
    return x[1:len(x)]

def force(x):
    if ((float(x[2]) + 360)%360) >= 180:
        FxCos = (float(x[0]) * float(x[3]))
        FxSin = (float(x[0]) * float(x[4]))
        FyCos = (float(x[1]) * float(-x[3]))
        FySin = (float(x[1]) * float(x[4]))
        x['FX'] = FxCos + FySin
        x['FY'] = FxSin + FyCos
        return x[['FX','FY']]
    else:
        FxCos = (float(x[0]) * float(x[3]))
        FxSin = (float(x[0]) * float(-x[4]))
        FyCos = (float(x[1]) * float(x[3]))
        FySin = (float(x[1]) * float(x[4]))
        x['FX'] = FxCos + FySin
        x['FY'] = FxSin + FyCos
        return x[['FX','FY']]
def GetPDMS():
    root.filename = filedialog.askopenfilename(initialdir='C:/',title='Choose your PDMS report file', filetypes=[('CSV','*.csv')])
    PDMSEntry.insert(0,root.filename)
def GetAutoPipe():
    root.filename = filedialog.askopenfilename(initialdir='C:/',title='Choose your AutoPipe report file', filetypes=[('XLS','*.xlsx')])
    AutoPipeEntry.insert(0,root.filename)
def GetFolderDir():
    root.filename = filedialog.askdirectory(initialdir='C:/',title='Choose folder Directory',)
    PathEntry.insert(0,root.filename + '/')
if __name__ == '__main__':
    def startProgram():
        Path = PathEntry.get()
        PDMS = PDMSEntry.get()
        AutoPipe = AutoPipeEntry.get()
        Com = [Com1.get(),Com2.get(),Com3.get(),Com4.get(),Com5.get(),Com6.get(),Com7.get(),Com8.get(),Com9.get(),Com10.get(),Com11.get(),Com12.get(),Com13.get(),Com14.get()]
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
        #Zwracanie wartości bezwzględej
        AutoPipeFile['ABS(FX)'] = [np.absolute(x) for x in AutoPipeFile['FX']]
        AutoPipeFile['ABS(FY)'] = [np.absolute(x) for x in AutoPipeFile['FY']]
        AutoPipeFile['ABS(FZ)'] = [np.absolute(x) for x in AutoPipeFile['FZ']]
        #Wybór maxymalnej wartości dla każdej osi
        AutoPipeFileFx = AutoPipeFile[['NAME','ABS(FX)']].groupby('NAME').max().reset_index()
        AutoPipeFileFx = pd.merge(AutoPipeFileFx,AutoPipeFile[['FX','Combination','ABS(FX)']],on='ABS(FX)',how='left').rename(columns={'Combination':'CombinationFx'})
        AutoPipeFileFy = AutoPipeFile[['NAME','ABS(FY)']].groupby('NAME').max().reset_index()
        AutoPipeFileFy = pd.merge(AutoPipeFileFy,AutoPipeFile[['FY','Combination','ABS(FY)']],on='ABS(FY)',how='left').rename(columns={'Combination':'CombinationFy'})
        AutoPipeFileFz = AutoPipeFile[['NAME','ABS(FZ)']].groupby('NAME').max().reset_index()
        AutoPipeFileFz = pd.merge(AutoPipeFileFz,AutoPipeFile[['FZ','Combination','ABS(FZ)']],on='ABS(FZ)',how='left').rename(columns={'Combination':'CombinationFz'})
        #Łączenie wszystkich osi razem do jednego pliku
        allDf = [AutoPipeFileFx,AutoPipeFileFy,AutoPipeFileFz]
        MergedData = reduce(lambda x,y: pd.merge(x,y,on=['NAME'],how='outer'),allDf)

        #Łączenie danych z PDMS'a i AutoPipe'a
        FinalReport = pd.merge(MergedData,PdmsFile[['NAME','ORIANGLE','SIN','COS','Description']],on='NAME',how='left')
        FinalReport[['FA','FL']] = FinalReport[['FX','FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
        FinalReport['FV'] = FinalReport['FZ']
        FinalReport = FinalReport[['NAME','Description','FX','CombinationFx','FY','CombinationFy','FZ','CombinationFz','FA','FL','FV']]

        # Wygenerowanie raportu końcowego
        FinalReport.to_excel(f'{Path}{Out}')
        #Wiadomość na koniec generowania raportu
        messagebox.showinfo('Raport Gotowy!',f'Plik został zapisany pod scieżką: {Path}{Out}')

    #Tkinter start
    root = Tk()
    root.title('Report Creator')
    ImgPath = pathlib.Path(__file__).parent.absolute()
    root.iconbitmap(f'{ImgPath}\\img\\logo.ico') 
    root.geometry('550x300')
   

    #Nazwa raportu PDMS
    PDMSEntry = Entry(root, width=55,borderwidth=2)
    PDMSEntry.grid(column=1,row=1)
    PdmsButton = Button(root, text='Get PDMS report directory',command=GetPDMS,width=25).grid(column=0,row=1)
    #Nazwa raportu Autopipe
    AutoPipeEntry = Entry(root, width=55,borderwidth=2)
    AutoPipeEntry.grid(column=1,row=2)
    AutoPipeButton = Button(root, text='Get AutoPipe report directory',command=GetAutoPipe,width=25).grid(column=0,row=2)
    #Scieżka do plików
    PathEntry = Entry(root, width=55,borderwidth=2)
    PathEntry.grid(column=1,row=0)
    PathButton = Button(root, text='Get folder directory',command=GetFolderDir,width=25).grid(column=0,row=0)


    #Kombinacja do wybrania
    Com1 = StringVar()
    Com2 = StringVar()
    Com3 = StringVar()
    Com4 = StringVar()
    Com5 = StringVar()
    Com6 = StringVar()
    Com7 = StringVar()
    Com8 = StringVar()
    Com9 = StringVar()
    Com10 = StringVar()
    Com11 = StringVar()
    Com12 = StringVar()
    Com13 = StringVar()
    Com14 = StringVar()
    Combinations = [('Gravity{1}',Com1),
                    ('GT1P1{2}',Com2),
                    ('GT1P1W1{2}',Com3),
                    ('GT1P1W2{2}',Com4),
                    ('GT1P1W3{2}',Com5),
                    ('GT1P1W4{2}',Com6),
                    ('GT1P1U1{4}',Com7),
                    ('User 3{5}',Com8),
                    ('GT1P1U3{5}',Com9),
                    ('GT2P2{6}',Com10),
                    ('Hydrotest-NL{7}',Com11),
                    ('MAX1',Com12),
                    ('MIN1',Com13),
                    ('EXTREME-OCCASIONAL',Com14)
                    ]
    i= 0
    for text,com in Combinations:
        if i <=6:
            Checkbutton(root,text=text,variable=com, onvalue=text,offvalue='').grid(column=0,row=i+3)
        else:
            Checkbutton(root,text=text,variable=com, onvalue=text,offvalue='').grid(column=1,row=(i-6)+2)
        i += 1
    StartButton = Button(root,text='Start',command=startProgram,height=1,width=10).grid(column=0,columnspan=2)
    root.mainloop()



