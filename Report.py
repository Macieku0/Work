import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re

def rangeLetter(cells):
    if not re.search(r'\d',cells):
        range_letter = cells.upper()
    else:
        range_letter = cells[:re.search(r'\d',cells).start()].upper()
    return range_letter


def rangeNumber(cells):
    range_number = cells[re.search(r'\d',cells).start():]
    return range_number


def colNameToColNumber(col_name):
    col_name = rangeLetter(col_name)
    if len(col_name) == 1:
        col_number = ord(col_name[0]) - 64
    elif len(col_name) == 2:
        col_number = ord(col_name[1]) - 64
        col_number += (ord(col_name[0]) - 64) * 26
    elif len(col_name) == 3:
        col_number = ord(col_name[2]) - 64
        col_number += (ord(col_name[1]) - 65) * 26
        col_number += (ord(col_name[0]) - 64) * 702
    return col_number

def colNumberToColLetter(col_number):
    if col_number < 27:
        col_name = chr(col_number+64)
    elif col_number < 703:
        col_name = chr((col_number//26)+64) + chr((col_number % 26)+64)
    elif col_number <= 703:
        col_name =  chr((col_number//702)+64) + chr(((col_number-702)//26)+65) + chr((col_number % 26)+64)
    return col_name


def adjustRangeFormat(start_cell,src,writer,sheet,index):
    length = src.shape[0]
    width = src.shape[1]
    end_cell = colNumberToColLetter(colNameToColNumber(start_cell)+width-index) + str(int(rangeNumber(start_cell)) + length)

    format_text = writer.book.add_format()
    column_align = writer.book.add_format({'align':'center'})
    column_wrap = writer.book.add_format({'align':'center','text_wrap':True})
    format_text.set_bottom()
    format_text.set_top()
    format_text.set_left()
    format_text.set_right()
    writer.sheets[sheet].conditional_format(f'{start_cell}:{end_cell}',{'type':'cell','criteria':'>=','value':0,'format':format_text})
    writer.sheets[sheet].conditional_format(f'{start_cell}:{end_cell}',{'type':'cell','criteria':'<','value':0,'format':format_text})

    for col in src:
        src2 = src.copy()
        src2[f'{col}'] = [str(x) for x in src2[f'{col}']]
        if len(col) > src2[f'{col}'].str.len().max():
            width = len(col)
        else:
            width = src2[f'{col}'].str.len().max()

        col_index = src.columns.get_loc(f'{col}') + colNameToColNumber(start_cell) - index

        if width <= 45:
            writer.sheets[sheet].set_column(col_index,col_index,width +1,column_align)
        else:
            writer.sheets[sheet].set_column(col_index,col_index,width +1,column_wrap)

    return [start_cell,end_cell]

def countProgress(src,zone):
    src = src[src['Zone'] == zone]
    src['Progress %'] = src['Value'] * src['Progress'] / src['Value'].sum()
    sumall = '{:.2%}'.format(src['Progress %'].sum())
    src['Progress %'] = src['Progress %'].astype(float).map('{:.2%}'.format)
    return src, sumall

def template(zone):
    src = pd.DataFrame({'Zone':[zone] * 9,'Status':['unset','P1','P2','P3','P4','P5','P6','P7','P8'],'Progress':[0,0.2,0.25,0.35,0.45,0.65,0.85,0.95,1],'Value': [0]*9,'Progress %': ['{:.2%}'.format(0)]*9})
    return src

def removeHelp(src):
    help_del = src.copy()
    help_del['without_help'] = help_del.apply(lambda x: True if ('HELP' in x['Zone'] or 'ACCESS' in x['Zone'] or 'REVIEW' in x['Zone']) else False,axis=1)
    help_del = help_del[help_del['without_help'] == False]
    help_del = help_del[['Zone','Name','Type','Status','User','Designer']]
    return help_del

def create_chart(writer,workbook,sheet,cells):
    chart = workbook.add_chart({'type':'column'})
    by_series = False
    if by_series:
        end_column = colNumberToColLetter(colNameToColNumber(cells[1]) + 2)
        start_row = int(rangeNumber(cells[0])) + 1
        end_row = int(rangeNumber(cells[1]))
        for x in range(start_row,end_row):
            chart.add_series({
                'values':f'={sheet}!$D${x}:$D${x}',
                'categories':f'={sheet}!$B${x}:$B${x}',
                'name':f'={sheet}!$B${x}:$B${x}'
            })
        chart.set_title({'name':f'={sheet}!$A${start_row}:$A${start_row}'})
        chart.set_table()
        writer.sheets[sheet].insert_chart(f'{end_column}{start_row}',chart)
    else:
        end_column = colNumberToColLetter(colNameToColNumber(cells[1]) + 2)
        start_row = int(rangeNumber(cells[0])) + 1
        end_row = int(rangeNumber(cells[1]))
        chart.add_series({
            'values':f'={sheet}!$D${start_row}:$D${end_row}',
            'categories':f'={sheet}!$B${start_row}:$B${end_row}',
            'name':f'={sheet}!$D${start_row-1}:$D${start_row-1}'
        })
        chart.set_title({'name':f'={sheet}!$A${start_row}:$A${start_row}'})
        chart.set_table()
        writer.sheets[sheet].insert_chart(f'{end_column}{start_row}',chart)

if __name__ == '__main__':

    progress = pd.DataFrame({'Status':['unset','P1','P2','P3','P4','P5','P6','P7','P8'],'Progress':[0,0.2,0.25,0.35,0.45,0.65,0.85,0.95,1]})

    path = 'C:\\Dev\\work\Work_python\\Statusy_190018_USER.csv'
    src1 = pd.read_csv(path,sep='|',decimal='.')
    src1 = src1.rename(columns={'NAME OF ZONE':'Zone','NAME':'Name','TYPE':'Type',':STATUS':'Status','userm':'User',':DESNAME':'Designer'})

    writer = pd.ExcelWriter('C:\\Dev\\work\\Work_python\\test.xlsx',engine='xlsxwriter')
    workbook = writer.book
    column_align = writer.book.add_format({'align':'center'})
    column_wrap = writer.book.add_format({'align':'center','text_wrap':True})

    src1 = removeHelp(src1)
    src1['Zone'] = [x[1:x.find('.')] for x in src1['Zone']]
    pipe = src1[src1['Type'] == 'PIPE']
    equip = src1[src1['Type'] == 'EQUI']
    pipe = pipe.reset_index(drop=True)

    pipe.to_excel(writer,sheet_name='Pipes')
    adjustRangeFormat('A1',pipe,writer,'Pipes',0)

    pipe = pipe[['Zone','Status']].value_counts()
    pipe = pipe.to_frame()
    pipe = pipe.rename(columns={0:'Value'})
    pipe['Value'] = pipe['Value'].astype(int)
    pipe = pipe.reset_index(level=[1,0])
    pipe = pipe.merge(progress,on='Status',how='left')

    worksheet = workbook.add_worksheet('Status')
    writer.sheets['Status'] = worksheet

    i = 1
    lista_zon = pipe['Zone'].unique()
    for zone in lista_zon:
        template_sheet = template(zone)
        globals()[f'Z_{zone}'],globals()[f'Z_{zone}_suma'] = countProgress(pipe,zone)
        globals()[f'Z_{zone}'] = template_sheet.append(globals()[f'Z_{zone}']).groupby(['Zone','Status','Progress']).max().reset_index()

        cells = adjustRangeFormat(f'A{i+1}',globals()[f'Z_{zone}'],writer,'Status',1)
        
        # print(chart_data)
        create_chart(writer,workbook,'Status',cells)

        worksheet.write(f'D{i+11}','Suma')
        worksheet.write(f'E{i+11}',globals()[f'Z_{zone}_suma'])
        format_text = writer.book.add_format()
        format_text.set_bottom()
        format_text.set_top()
        format_text.set_left()
        format_text.set_right()
        worksheet.conditional_format(f'D{i+11}:E{i+11}',{'type':'cell','criteria':'>=','value':0,'format':format_text})
        worksheet.conditional_format(f'D{i+11}:E{i+11}',{'type':'cell','criteria':'<','value':0,'format':format_text})

        globals()[f'Z_{zone}'].to_excel(writer,sheet_name='Status',startrow = i, startcol = 0, index=False)
        i +=12

    writer.save()