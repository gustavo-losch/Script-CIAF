from dbfread import DBF
import pandas as pd
import openpyxl as xl
import sys
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, BaseDocTemplate, PageTemplate, Frame, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

# Abrindo arquivo DBF e criando o DataFrame
path = ''r'C:\Users\Oficina\Desktop\Gustavo\Repositórios\Script-CIAF\Conversão\ORDEM-SERV.DBF'
dbf = DBF(path)
dataResult = pd.DataFrame(iter(dbf))

#Filtro de colunas e transformando em datetime as datas
tabela = dataResult[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL', 'CODCLI', 'NOMECLIENT']]
tabela.loc[:, 'DATAFECHA'] = pd.to_datetime(tabela['DATAFECHA'], errors='coerce')

datai = "2024-01-30"#input()
dataf = "2024-03-30"#input()

#Inputs dos intervalos de data para aplicar filtro
data_inicial = pd.to_datetime(datai)
data_final = pd.to_datetime(dataf)

#Filtros de cliente e data
df_filtrado = tabela[(tabela['DATAFECHA'] >= data_inicial) & (tabela['DATAFECHA'] <= data_final)]
df = df_filtrado[df_filtrado['CODCLI'].str.contains('147')]

#Correção de dados e soma dos valores
df['DATAFECHA'] = df['DATAFECHA'].fillna(pd.Timestamp('2024-01-01'))
df['DATAFECHA'] = df['DATAFECHA'].dt.strftime('%d/%m/%Y')
df['VL_TOTAL'] = df['VL_TOTAL'].round(0).astype(int).astype(str) + ',00'
df['NR_ORDEM'] = df['NR_ORDEM'].round(0).astype(int).astype(str)
total = df['VL_TOTAL'].str.replace(',00', '').astype(int).sum()
df.loc['Total'] = pd.Series(str(total) + ',00', index = ['VL_TOTAL'])
df.loc['Total'] = df.loc['Total'].fillna(' ')
df.loc['Total', 'OS CIAF'] = 'Total'

#Renomeação das colunas
df['OS Bergerson'] = df['NRSERIE'].str[-4:]
df['OS CIAF'] = df['NR_ORDEM']
df['Valor'] = df['VL_TOTAL']
df['Data de Fechamento'] = df['DATAFECHA']
df = df[['OS CIAF', 'OS Bergerson', 'Data de Fechamento', 'Valor']]

#Criação do PDF com reportlab
styles = getSampleStyleSheet()

#Criação dos diferentes estilos utilizados nos documentos
title_style = ParagraphStyle(
    'Title',
    parent=styles['Title'],
    fontSize=14,
    textColor=colors.black,
    alignment=1,  # 1 = TA_CENTER
)

subtitle_style = ParagraphStyle(
    'Subtitle',
    parent=styles['Normal'],
    fontsize=10,
    textColor=colors.black,
    alignment=1,
)

footer_style = ParagraphStyle(
    'Footer',
    parent=styles['Normal'],
    fontSize=10,
    textColor=colors.grey,
    alignment=2,  # 2 = TA_RIGHT
)

table_style = TableStyle([
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')
    ])

#Conteúdo do PDF
title_text = "RDA Design"
title = Paragraph(title_text, title_style)

date_range_text = "De {} até {}".format(data_inicial.strftime('%d/%m/%Y'), data_final.strftime('%d/%m/%Y'))
date_range = Paragraph(date_range_text, subtitle_style)

footer_text = "Cesar Augusto Ribeiro do Amaral"
footer = Paragraph(footer_text, footer_style)

data = [df.columns.to_list()] + df.values.tolist()
table = Table(data)
table.setStyle(table_style)

text_below_table = "Tabela de preços dos serviços prestados."
text_paragraph = Paragraph(text_below_table, subtitle_style)

#Criação do arquivo e build dos elementos
doc = SimpleDocTemplate("Tabela Bergerson.pdf", pagesize=letter)

elements = [title, date_range, Spacer(1,35) , table, Spacer(1,15),text_paragraph, Spacer(1,30), footer]
doc.build(elements)