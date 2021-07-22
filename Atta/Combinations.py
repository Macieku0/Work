import pandas as pd



#Listowanie kombinacji dla danego raportu
def CombinationsList(path):
    #Wczytanie pliku autopipe
    AutoPipeFile = pd.read_excel(f'{path}')
    #Selekcja kolumn
    AutoPipeFile = AutoPipeFile['Combination'].drop_duplicates()
    AutoPipeFile = AutoPipeFile.dropna()
    comList = list(AutoPipeFile)
    return comList
