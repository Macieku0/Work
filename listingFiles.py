import os

i = 0
lista = ['']
# rozszerzenie plików które mają być skopiowane
extension = '.pdf'
# Scieżka z której mają być kopiowanie pliki
dirPathFrom = "G:/#2019/190018/M/EXPORT/"
# Scieżka do której mają być kopiowane pliki
dirPathTo = 'G:/#2019/190018/M/POMOC/MU/PDFY/'

#Scieżka pod którą zostanie zapisany wykaz skopiowanych plików
fullList = open(f'{dirPathTo}/wykaz_plikow.txt', 'w')

#Funkcja zwracająca tylko pliki z odpowiednim rozszerzeniem
def createList (dir, ext):
    return (file for file in os.listdir(dir) if file.endswith(f'{ext}'))

#Funckja główna
def main():
    for root,dirs,files in os.walk(dirPathFrom):
        pdfs = createList(os.path.join(root),extension)
        for pdf in pdfs:
            i += 1
            entry =  f'{i}. {pdf} z folderu {os.path.join(root)}\n' 
            lista.append(entry)

#Wiadomość końcowa
print(f'Skończone, skopiowano {i},plików z rozszerzeniem {extension}')

#Zapis wykazu
fullList.writelines(lista)
fullList.close()
