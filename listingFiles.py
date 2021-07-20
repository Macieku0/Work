import os

lista = ['']
# rozszerzenie plików które mają być skopiowane
extension = '.pdf'
# Scieżka z której mają być kopiowanie pliki
dirPathFrom = "G:/#2019/190018/M/PT/UDT/PED/dokumentacja rejestracyjna/ZDT/"
# dirPathFrom = "C:\\Users\\macie\\Pulpit\\III_CAT_SPRAWDZENIE_SRUB\\20210611_REGISTRATION_UDT"
# Scieżka do której mają być kopiowane pliki
dirPathTo = 'G:/#2019/190018/M/PT/UDT/PED/dokumentacja rejestracyjna/ZDT/'
# dirPathTo = "C:\\Users\\macie\\Pulpit\\III_CAT_SPRAWDZENIE_SRUB\\20210611_REGISTRATION_UDT"

#Scieżka pod którą zostanie zapisany wykaz skopiowanych plików
fullList = open(f'{dirPathTo}/wykaz_plikow.txt', 'w')

#Funkcja zwracająca tylko pliki z odpowiednim rozszerzeniem
def createList (dir, ext):
    return (file for file in os.listdir(dir) if file.endswith(f'{ext}'))

#Funckja główna
def main():
    global i
    i = 0
    for root,dirs,files in os.walk(dirPathFrom):
        pdfs = createList(os.path.join(root),extension)
        for pdf in pdfs:
            i += 1
            entry =  f'{i}|{pdf}|{os.path.join(root)}\n' 
            lista.append(entry)

#Zapis wykazu
main()
fullList.writelines(lista)
fullList.close()
