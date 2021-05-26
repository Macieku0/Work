from numpy.core.fromnumeric import sort
import pandas as pd
import numpy as np
import math


#Nazwa raportu PDMS
print("Wpisz nazwę pliku z PDMS'a wraz z rozszerzeniem np.'zrodlo.csv'")
PDMS = f'{input()}'
#Nazwa raportu Autopipe
print("Wpisz nazwę pliku z AutoPipe wraz z rozszerzeniem np.'plik.xlsx'")
AutoPipe = f'{input()}'
#Nazwa raportu wyjściowego
Out = 'Raport_policzony.xlsx'
#Scieżka do plików
print("Wpisz ścieżkę do pliku 'C:/folder_z_plikami/'")
Path = f'{input()}'
#Kombinacja do wybrania
print("Wpisz kombinację case'a np.'GT1P1U3{5}'")
Com = f'{input()}'

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

if __name__ == '__main__':
    #'Tag No.'  jest równie numerowi zamocowania
    PdmsFile = pd.read_csv(f'{Path}{PDMS}',sep='|',decimal='.')
    PdmsFile = PdmsFile.replace("=","'=",regex=True)
    PdmsFile = PdmsFile.replace('degree','',regex=True)
    PdmsFile['ORIANGLE'] = [clean_angles(oriangle) for oriangle in PdmsFile['ORIANGLE']]
    PdmsFile['NAME'] = [clean_name(name) for name in PdmsFile['NAME']]
    PdmsFile['SIN'] = [sin(oriangle) for oriangle in PdmsFile['ORIANGLE']]
    PdmsFile['COS'] = [cos(oriangle) for oriangle in PdmsFile['ORIANGLE']]

    AutoPipeFile = pd.read_excel(f'{Path}{AutoPipe}')
    AutoPipeFile = AutoPipeFile[['Tag No.','Combination','GlobalFX','GlobalFY','GlobalFZ']]
    AutoPipeFile = AutoPipeFile.drop(0)
    AutoPipeFile = AutoPipeFile[AutoPipeFile['Combination'] == Com]
    AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']] = AutoPipeFile[['GlobalFX','GlobalFY','GlobalFZ']].astype(float)
    AutoPipeFile.rename(columns={'Tag No.':'NAME','GlobalFZ':'FZ','GlobalFX':'FX','GlobalFY':'FY'},inplace=True)
    AutoPipeFile['NAME'] = [clean_name(name) for name in AutoPipeFile['NAME']]
    AutoPipeFile = AutoPipeFile[['NAME','FX','FY','FZ']].groupby('NAME').sum()

    FinalReport = pd.merge(AutoPipeFile,PdmsFile[['NAME','ORIANGLE','SIN','COS']],on='NAME',how='left')
    FinalReport[['FA','FL']] = FinalReport[['FX','FY','ORIANGLE','COS','SIN',]].apply(force,axis=1)
    FinalReport['FV'] = FinalReport['FZ']
    FinalReport = FinalReport[['NAME','FX','FY','FZ','FA','FL','FV']]

    # Wygenerowanie raportu końcowego
    print(f'Wszystko poszło "ok" plik zostanie zapisany pod scieżką: {Path}')
    FinalReport.to_excel(f'{Path}{Out}')
    print('Plik zapisany')
    print('Wciśnij enter aby zakończyć')
    koniec = input()

