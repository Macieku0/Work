import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st


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


if __name__ == '__main__':

    progress = pd.DataFrame({'Status':['unset','P1','P2','P3','P4','P5','P6','P7','P8'],'Progress':[0,0.2,0.25,0.35,0.45,0.65,0.85,0.95,1]})

    path = 'C:\\Dev\\work\Work_python\\Statusy_190018_USER.csv'
    src1 = pd.read_csv(path,sep='|',decimal='.')
    src1 = src1.rename(columns={'NAME OF ZONE':'Zone','NAME':'Name','TYPE':'Type',':STATUS':'Status','userm':'User',':DESNAME':'Designer'})

    src1 = removeHelp(src1)
    src1['Zone'] = [x[1:x.find('.')] for x in src1['Zone']]
    pipe = src1[src1['Type'] == 'PIPE']
    equip = src1[src1['Type'] == 'EQUI']
    pipe = pipe.reset_index(drop=True)


    pipe = pipe[['Zone','Status']].value_counts()
    pipe = pipe.to_frame()
    pipe = pipe.rename(columns={0:'Value'})
    pipe['Value'] = pipe['Value'].astype(int)
    pipe = pipe.reset_index(level=[1,0])
    pipe = pipe.merge(progress,on='Status',how='left')

    st.title('Progress report')

    st.sidebar.header('User input features')
    

    i = 1
    lista_zon = pipe['Zone'].unique()
    wybranazona = st.sidebar.multiselect("Zony",lista_zon)
    for zone in lista_zon:
        template_sheet = template(zone)
        globals()[f'Z_{zone}'],globals()[f'Z_{zone}_suma'] = countProgress(pipe,zone)
        globals()[f'Z_{zone}'] = template_sheet.append(globals()[f'Z_{zone}']).groupby(['Zone','Status','Progress']).max().reset_index()
        for zona in wybranazona:
            if zona == zone:
                st.write(globals()[f'Z_{zone}'])
                st.text(f'Suma progresu dla zony {zone} wynosi {globals()[f"Z_{zone}_suma"]}')
                globals()[f'Z_{zone}'] = globals()[f'Z_{zone}'][['Status','Value']].set_index('Status')
                chart_data = pd.DataFrame(
                globals()[f'Z_{zone}'])

                st.bar_chart(chart_data)
            i +=12

    
    if st.checkbox('Zbiorcze'):
        variables = st.multiselect("Select zone/zones",list(lista_zon))
        chart_data = pipe.copy()
        st.write(chart_data[(chart_data['Zone'].isin(variables))][['Zone','Status','Value']])
        chart_data = chart_data[['Status','Value','Zone']].groupby(['Status','Zone']).sum().reset_index()
        chart_data = chart_data[(chart_data['Zone'].isin(variables))]
        chart_data = chart_data.pivot(index='Status',columns='Zone',values='Value')
        chart_data = chart_data.rename_axis(None,axis=1)
        chart_data = chart_data.replace(to_replace=np.nan, value=0)
        st.bar_chart(chart_data)