
import pandas as pd
import re

path = "C:\\Users\\macie\Pulpit\\III_CAT_SPRAWDZENIE_SRUB\\20210601_BOLTS\\BOLT.TXT"

# src = src[['PIPELINE-REFERENCE','ITEM-CODE','QNTY','DESCRIPTION']].rename(columns={'QNTY':'QUANTITY','PIPELINE-REFERENCE':'PIPELINE NAME'})
# src[['SECTION','PN']] = ['ŚRUBY, NAKRĘTKI','-']
# src['DN1'] = [x[:3] for x in src['ITEM-CODE']]
# src['PN'] = '-'
# src
# print(src)
with open(path, "r") as file:
    lines = file.readlines()
    lista = []
    global newList
    newList = ['','','','','']
    i = 0
    for line in lines:
        if not line.isspace():
            a = str(line)
            if (len(a) <= 111 and len(a) >= 106):
                description = a[2:40]
                itemCode = a[70:92]
                quantity = re.sub('\n','',a[108:111])
            if (len(a) <= 50 and len(a) >= 10):
                if a[1:9] == 'PIPELINE':
                    newList = ['','','','','']
                    pipeline = re.sub('\n','',a[14:len(a)])
                    newList[0] = pipeline
                elif a[5:7] == a.split()[0][0:2]:
                    secondDesc = re.sub('\n','',a[5:len(a)])
                    description = description + ', ' + secondDesc
                    newList[1] = description
                    newList[2]= itemCode
                    newList[3] = quantity
                    newList[4] =  secondDesc
                    przepis = newList.copy()
                    i += 1
                    lista.append(przepis)
            if (len(a) <= 10 and len(a) >= 0):
                material = re.sub('\n','',a[5:len(a)])
                newList[1] = description + ', ' + material
                newList[4] = material
                del lista[i-1]
                przepis = newList.copy()
                lista.append(przepis)
        # if i == 5:
        #     break

src = pd.DataFrame(lista,columns=['PIPLINE NAME','DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL'])
src[['SECTION','PN']] = ['ŚRUBY, NAKRĘTKI','-']
src['DN1'] = [x[:3] for x in src['ITEM-CODE']]


poRurach = src[['PIPLINE NAME','ITEM-CODE','DESCRIPTION','MATERIAL']].copy()
print(poRurach)
poRurach['QUANTITY'] = [int(x) for x in poRurach['QUANTITY']]
poRurach = poRurach.groupby([['PIPLINE NAME','ITEM-CODE','DESCRIPTION','MATERIAL']]).sum().reset_index()
zbiorowe = src.copy()
zbiorowe['QUANTITY'] = [int(x) for x in zbiorowe['QUANTITY']]
zbiorowe = zbiorowe[['DESCRIPTION','ITEM-CODE','QUANTITY','MATERIAL']].groupby([['ITEM-CODE','DESCRIPTION','MATERIAL']]).sum().reset_index()
writer = pd.ExcelWriter("C:\\Users\\macie\Pulpit\\III_CAT_SPRAWDZENIE_SRUB\\20210601_BOLTS\\BOLT.xlsx")
poRurach.to_excel(writer, sheet_name='Po rurociągach')
zbiorowe.to_excel(writer, sheet_name='Zbiorowe')

writer.save()