import os
import shutil
i = 0
lista = ['']
# rozszerzenie plików które mają być skopiowane
extension = 'pdf'
# Scieżka z której mają być kopiowanie pliki
dirPathFrom = "G:/#2019/190018/M/EXPORT/"
# Scieżka do której mają być kopiowane pliki
dirPathTo = 'G:/#2019/190018/M/POMOC/MU/PDFY/'

#Lista do sprawdzenia
checkList = ['TKIS-PIP-3120-MB-00011-03.pdf','TKIS-PIP-3120-MB-00015-05.pdf','TKIS-PIP-3120-MB-00541-01.pdf','TKIS-PIP-3120-MB-00020-02.pdf','TKIS-PIP-3120-MB-00023-04.pdf','TKIS-PIP-3120-MB-00027-05.pdf','TKIS-PIP-3120-MB-00031-03.pdf','TKIS-PIP-3120-MB-00033-04.pdf','TKIS-PIP-3120-MB-00037-03.pdf','TKIS-PIP-3120-MB-00042-02.pdf','TKIS-PIP-3120-MB-00045-03.pdf','TKIS-PIP-3120-MB-00059-03.pdf','TKIS-PIP-3120-MB-00061-03.pdf','TKIS-PIP-3120-MB-00062-05.pdf','TKIS-PIP-3120-MB-00063-03.pdf','TKIS-PIP-3120-MB-00068-04.pdf','TKIS-PIP-3120-MB-00082-02.pdf','TKIS-PIP-3120-MB-00091-04.pdf','TKIS-PIP-3120-MB-00110-01.pdf','TKIS-PIP-3120-MB-00113-02.pdf','TKIS-PIP-3120-MB-00114-02.pdf','TKIS-PIP-3120-MB-00125-02.pdf','TKIS-PIP-3120-MB-00542-01.pdf','TKIS-PIP-3120-MB-00134-03.pdf','TKIS-PIP-3120-MB-00135-02.pdf','TKIS-PIP-3120-MB-00136-03.pdf','TKIS-PIP-3120-MB-00137-03.pdf','TKIS-PIP-3120-MB-00138-02.pdf','TKIS-PIP-3120-MB-00139-03.pdf','TKIS-PIP-3120-MB-00140-06.pdf','TKIS-PIP-3120-MB-00141-02.pdf','TKIS-PIP-3120-MB-00142-02.pdf','TKIS-PIP-3120-MB-00143-02.pdf','TKIS-PIP-3120-MB-00144-05.pdf','TKIS-PIP-3120-MB-00146-02.pdf','TKIS-PIP-3120-MB-00147-02.pdf','TKIS-PIP-3120-MB-00148-01.pdf','TKIS-PIP-3120-MB-00145-04.pdf','TKIS-PIP-3120-MB-00151-03.pdf','TKIS-PIP-3120-MB-00153-02.pdf','TKIS-PIP-3120-MB-00505-02.pdf','TKIS-PIP-3120-MB-00506-01.pdf','TKIS-PIP-3120-MB-00507-04.pdf','TKIS-PIP-3120-MB-00508-03.pdf','TKIS-PIP-3120-MB-00164-02.pdf','TKIS-PIP-3120-MB-00165-01.pdf','TKIS-PIP-3120-MB-00166-02.pdf','TKIS-PIP-3120-MB-00167-02.pdf','TKIS-PIP-3120-MB-00168-06.pdf','TKIS-PIP-3120-MB-00172-02.pdf','TKIS-PIP-3120-MB-00174-05.pdf','TKIS-PIP-3120-MB-00176-06.pdf','TKIS-PIP-3120-MB-00177-04.pdf','TKIS-PIP-3120-MB-00180-03.pdf','TKIS-PIP-3120-MB-00182-04.pdf','TKIS-PIP-3120-MB-00186-03.pdf','TKIS-PIP-3120-MB-00204-03.pdf','TKIS-PIP-3120-MB-00231-02.pdf','TKIS-PIP-3120-MB-00237-04.pdf','TKIS-PIP-3120-MB-00241-03.pdf','TKIS-PIP-3120-MB-00242-04.pdf','TKIS-PIP-3130-MB-00014-02.pdf','TKIS-PIP-3130-MB-00015-03.pdf','TKIS-PIP-3130-MB-00017-02.pdf','TKIS-PIP-3130-MB-00018-03.pdf','TKIS-PIP-3130-MB-00030-02.pdf','TKIS-PIP-3130-MB-00023-04.pdf','TKIS-PIP-3130-MB-00032-02.pdf','TKIS-PIP-3130-MB-00033-02.pdf','TKIS-PIP-3130-MB-00034-03.pdf','TKIS-PIP-3130-MB-00027-03.pdf','TKIS-PIP-3130-MB-00036-04.pdf','TKIS-PIP-3130-MB-00038-02.pdf','TKIS-PIP-3130-MB-00039-04.pdf','TKIS-PIP-3130-MB-00361-02.pdf','TKIS-PIP-3130-MB-00040-03.pdf','TKIS-PIP-3130-MB-00041-03.pdf','TKIS-PIP-3130-MB-00042-03.pdf','TKIS-PIP-3130-MB-00043-03.pdf','TKIS-PIP-3130-MB-00044-04.pdf','TKIS-PIP-3130-MB-00047-02.pdf','TKIS-PIP-3130-MB-00053-02.pdf','TKIS-PIP-3130-MB-00054-04.pdf','TKIS-PIP-3130-MB-00056-04.pdf','TKIS-PIP-3130-MB-00058-03.pdf','TKIS-PIP-3130-MB-00059-02.pdf','TKIS-PIP-3130-MB-00063-02.pdf','TKIS-PIP-3130-MB-00064-02.pdf','TKIS-PIP-3130-MB-00065-03.pdf','TKIS-PIP-3130-MB-00066-03.pdf','TKIS-PIP-3130-MB-00067-03.pdf','TKIS-PIP-3130-MB-00068-02.pdf','TKIS-PIP-3130-MB-00069-03.pdf','TKIS-PIP-3130-MB-00070-04.pdf','TKIS-PIP-3130-MB-00071-04.pdf','TKIS-PIP-3130-MB-00072-02.pdf','TKIS-PIP-3130-MB-00073-07.pdf','TKIS-PIP-3130-MB-00074-07.pdf','TKIS-PIP-3130-MB-00075-01.pdf','TKIS-PIP-3130-MB-00076-02.pdf','TKIS-PIP-3130-MB-00077-02.pdf','TKIS-PIP-3130-MB-00078-04.pdf','TKIS-PIP-3130-MB-00085-02.pdf','TKIS-PIP-3130-MB-00095-04.pdf','TKIS-PIP-3130-MB-00096-03.pdf','TKIS-PIP-3130-MB-00097-03.pdf','TKIS-PIP-3130-MB-00313-02.pdf','TKIS-PIP-3130-MB-00362-02.pdf','TKIS-PIP-3130-MB-00106-05.pdf','TKIS-PIP-3130-MB-00108-01.pdf','TKIS-PIP-3130-MB-00113-03.pdf','TKIS-PIP-3130-MB-00118-03.pdf','TKIS-PIP-3130-MB-00340-03.pdf','TKIS-PIP-3130-MB-00135-01.pdf','TKIS-PIP-3130-MB-00136-03.pdf','TKIS-PIP-3130-MB-00137-02.pdf','TKIS-PIP-3130-MB-00140-04.pdf','TKIS-PIP-3130-MB-00141-03.pdf','TKIS-PIP-3130-MB-00142-04.pdf','TKIS-PIP-3130-MB-00144-02.pdf','TKIS-PIP-3130-MB-00159-04.pdf','TKIS-PIP-3130-MB-00316-03.pdf','TKIS-PIP-0501-MB-00011-02.pdf','TKIS-PIP-0501-MB-00015-02.pdf','TKIS-PIP-0501-MB-00018-02.pdf','TKIS-PIP-0501-MB-00020-02.pdf','TKIS-PIP-0501-MB-00027-01.pdf','TKIS-PIP-0501-MB-00029-02.pdf','TKIS-PIP-3361-MB-00012-04.pdf','TKIS-PIP-3361-MB-00013-04.pdf','TKIS-PIP-3361-MB-00041-02.pdf','TKIS-PIP-3361-MB-00043-02.pdf','TKIS-PIP-3361-MB-00058-02.pdf','TKIS-PIP-3130-MB-00288-04.pdf','TKIS-PIP-3130-MB-00315-03.pdf']
#Scieżka pod którą zostanie zapisany wykaz skopiowanych plików
fullList = open(f'{dirPathTo}/wykaz_plikow.txt', 'w')

#Funkcja zwracająca tylko pliki z odpowiednim rozszerzeniem
def createList (dir, ext):
    return (file for file in os.listdir(dir) if file.endswith(f'.{ext}'))

#Funckja główna
for root,dirs,files in os.walk(dirPathFrom):
    pdfs = createList(os.path.join(root),extension)
    for pdf in pdfs:
        if pdf in checkList:
            shutil.copyfile(f'{os.path.join(root)}/{pdf}', f'{dirPathTo}/{pdf}')
            i += 1
            entry =  f'{i}. {pdf} z folderu {os.path.join(root)}\n' 
            lista.append(entry)

#Wiadomość końcowa
print(f'Skończone, skopiowano {i},plików z rozszerzeniem .{extension}')

#Zapis wykazu
fullList.writelines(lista)
fullList.close()
