#Listowanie kombinacji dla danego raportu

from os import supports_dir_fd
import pandas as pd

def CombinationsList(path):
    #Wczytanie pliku autopipe
    AutoPipeFile = pd.read_excel(f'{path}')
    #Selekcja kolumn
    AutoPipeFile = AutoPipeFile['Combination'].drop_duplicates()
    comList = list(AutoPipeFile)
    return comList

