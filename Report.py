import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

progress = pd.DataFrame({'Status':['unset','P1','P2','P3','P4','P5','P6','P7','P8'],'Progress':[0,0.2,0.25,0.35,0.45,0.65,0.85,0.95,1]})

path = 'C:\\Dev\\work\Work_python\\Statusy_190018_USER.csv'
src1 = pd.read_csv(path,sep='|',decimal='.')
src1 = src1.rename(columns={'NAME OF ZONE':'Zone','NAME':'Name','TYPE':'Type',':STATUS':'Status','userm':'User',':DESNAME':'Designer'})

def remove_help(src):
    help_del = src.copy()
    help_del['without_help'] = help_del.apply(lambda x: True if ('HELP' in x['Zone'] or 'ACCESS' in x['Zone'] or 'REVIEW' in x['Zone']) else False,axis=1)
    help_del = help_del[help_del['without_help'] == False]
    help_del = help_del[['Zone','Name','Type','Status','User','Designer']]
    return help_del

def count_progress(src,zone):
    src = src[src['Zone'] == zone]
    src['Progress %'] = src['Value'] * src['Progress'] / src['Value'].sum()
    sumall = '{:.2%}'.format(src['Progress %'].sum())
    src['Progress %'] = src['Progress %'].astype(float).map('{:.2%}'.format)
    return src, sumall




src1 = remove_help(src1)
src1['Zone'] = [x[1:x.find('.')] for x in src1['Zone']]
pipe = src1[src1['Type'] == 'PIPE']
equip = src1[src1['Type'] == 'EQUI']

pipe = pipe[['Zone','Status']].value_counts()
pipe = pipe.to_frame()
pipe = pipe.rename(columns={0:'Value'})
pipe['Value'] = pipe['Value'].astype(int)
pipe = pipe.reset_index(level=[1,0])
pipe = pipe.merge(progress,on='Status',how='left')


writer = pd.ExcelWriter('C:\\Dev\\work\\Work_python\\test.xlsx',engine='xlsxwriter')
workbook = writer.book
worksheet = workbook.add_worksheet('Status')
writer.sheets['Status'] = worksheet

i = 1
lista_zon = pipe['Zone'].unique()
for zona in lista_zon:
    globals()[f'Z_{zona}'],globals()[f'Z_{zona}_suma'] = count_progress(pipe,zona)
    globals()[f'Z_{zona}'].to_excel(writer,sheet_name='Status',startrow = 0, startcol = i)
    i +=7

writer.save()
# ax = pipe.T.plot(kind = 'bar')
# ylab = ax.set_ylabel('Values')