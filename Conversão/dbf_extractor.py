from dbfread import DBF
import pandas as pd
import openpyxl
import sys
import os

dbf = DBF(r'C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\Conversão\ORDEM-SERV.DBF')
dataResult = pd.DataFrame(iter(dbf))
tabela = dataResult[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL', 'CODCLI', 'NOMECLIENT']]
tabela.loc[:, 'DATAFECHA'] = pd.to_datetime(tabela['DATAFECHA'], errors='coerce')

datai = "2024-01-01"#input()
dataf = "2024-01-30"#input()

data_inicial = pd.to_datetime(datai)
data_final = pd.to_datetime(dataf)

df_filtrado = tabela[(tabela['DATAFECHA'] >= data_inicial) & (tabela['DATAFECHA'] <= data_final)]
df = df_filtrado[df_filtrado['CODCLI'].str.contains('147')]
df['DATAFECHA'] = df['DATAFECHA'].fillna(pd.Timestamp('2024-01-01'))
df['DATAFECHA'] = df['DATAFECHA'].dt.strftime('%d/%m/%Y')
df['VL_TOTAL'] = df['VL_TOTAL'].round(0).astype(int).astype(str) + ',00'
total = df['VL_TOTAL'].str.replace(',00', '').astype(int).sum()
df.loc['Total'] = pd.Series(str(total) + ',00', index = ['VL_TOTAL'])

df['OS Bergerson'] = df['NRSERIE'].str[-4:]
df['OS CIAF'] = df['NR_ORDEM']
df['Valor'] = df['VL_TOTAL']
df['Data de Fechamento'] = df['DATAFECHA']
df = df[['OS CIAF', 'OS Bergerson', 'Data de Fechamento', 'Valor']]

df.to_excel("Tabela Bergerson.xlsx", index=False)