from dbfread import DBF
import pandas as pd
import openpyxl

dbf = DBF(r'C:\Users\Oficina\Desktop\Gustavo\Repositórios\Script-CIAF\Conversão\ORDEM-SERV.DBF')
dataResult = pd.DataFrame(iter(dbf))
tabela = dataResult[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL', 'CODCLI', 'NOMECLIENT']]
df = tabela[tabela['CODCLI'].str.contains('147')]
df.loc[:, 'DATAFECHA'] = pd.to_datetime(df['DATAFECHA'], errors='coerce')

datai = input()
dataf = input()

data_inicial = pd.to_datetime(datai)
data_final = pd.to_datetime(dataf)

df_filtrado = df[(df['DATAFECHA'] >= data_inicial) & (df['DATAFECHA'] <= data_final)]

df_filtrado.to_excel("Tabela Bergerson2.xlsx")