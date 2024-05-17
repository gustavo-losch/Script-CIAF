from dbfread import DBF
import pandas as pd
import openpyxl

dbf = DBF(r'C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\Conversão\ORDEM-SERV.DBF')
dataResult = pd.DataFrame(iter(dbf))
tabela = dataResult[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL', 'CODCLI', 'NOMECLIENT']]
tabela.loc[:, 'DATAFECHA'] = pd.to_datetime(tabela['DATAFECHA'], errors='coerce')

datai = input()
dataf = input()

data_inicial = pd.to_datetime(datai)
data_final = pd.to_datetime(dataf)

df_filtrado = tabela[(tabela['DATAFECHA'] >= data_inicial) & (tabela['DATAFECHA'] <= data_final)]
df = df_filtrado[df_filtrado['CODCLI'].str.contains('147')]

df_filtrado = df[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL']]

print(df_filtrado)

#df_filtrado.to_excel("Tabela Bergerson2.xlsx")