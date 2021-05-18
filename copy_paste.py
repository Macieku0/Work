import os
import shutil
i = 0
lista = ['']
# rozszerzenie plików które mają być skopiowane
extension = 'pdf'
# Scieżka z której mają być kopiowanie pliki
dirPathFrom = "C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/20210420_FLANGE_CONNECTIONS_I_II_III_CAT/NOWE"
# Scieżka do której mają być kopiowane pliki
dirPathTo = 'C:/Users/macie/Pulpit/III_CAT_SPRAWDZENIE_SRUB/20210420_FLANGE_CONNECTIONS_I_II_III_CAT/Nowy folder'

fullList = open(dirPathTo + '/' + 'wykaz_plikow.txt', 'w')

def list (dir, ext):
    return (file for file in os.listdir(dir) if file.endswith('.' + ext))


for root,dirs,files in os.walk(dirPathFrom):
    pdfs = list(os.path.join(root),extension)
    for pdf in pdfs:
        shutil.copyfile(os.path.join(root) +'/'+pdf, dirPathTo + '/'+ pdf)
        i += 1
        entry =  str(i) + '. ' + pdf + 'z folderu ' + str(os.path.join(root)) + '\n' 
        lista.append(entry)

print('Skończone, skopiowano', i,'plików z rozszerzeniem .'+ extension)

fullList.writelines(lista)
fullList.close()
