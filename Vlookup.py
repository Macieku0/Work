import pandas as pd
import numpy as np
from pandas.core.indexes.base import Index

file_A = 'A.xlsx'
file_B = 'B.xlsx'
summary = 'summary.xlsx'

df_a = pd.read_excel(f'C:/Dev/Test/{file_A}')
df_b = pd.read_excel(f'C:/Dev/Test/{file_B}')

# print(df_a.head())
# print(df_b.head())

output = pd.merge(df_a,df_b[['Nr','Koszt']],on='Nr', how='left')
output['Różnica'] = output['Koszt_x'] - output['Koszt_y']


output = output.replace(np.nan,'',regex=True)

output.to_excel(f'C:/Dev/Test/{summary}', index=False)