import tabula
import pyautogui
import time as tm
from datetime import *
import tkinter
import customtkinter
from customtkinter import filedialog
import re
from dbfread import DBF
import pandas as pd
import openpyxl
import sys
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, BaseDocTemplate, PageTemplate, Frame, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Line
from PIL import Image, ImageTk
import shutil
import csv
from CTkPDFViewer import *
import pymupdf

def CRWindow():
    windowCR = customtkinter.CTkToplevel(principal)
    principal.iconify()
    windowCR.title("Contas a Receber")
    windowCR.geometry("400x250")
    windowCR.resizable(False, False)

    def iniciar_processo_cr():
        dataCR = entryCR.get()
        dataCR = re.sub('[/]', '', dataCR)
        if len(dataCR) == 8:
            print("Rodar")
            runCR.configure(text="O código irá rodar em 5 segundos.",  text_color="green")
            entryCR.configure(state="disabled")
            tm.sleep(5)

            tabela = tabula.read_pdf("teste2.pdf", pages="all")[0] #PDF Table read and transform to dataframe
            tabela.rename(columns=tabela.iloc[4], inplace = True)
            tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
            tabela = tabela.drop(0)
            tabela = tabela.drop(1)
            tabela = tabela.drop(2)
            tabela = tabela.drop(3)
            tabela = tabela.drop(4)
            ultimo = len(tabela) + 4
            tabela = tabela.drop(ultimo)

            pyautogui.click(975,1055) #Preparing CIAF to recieve the commands
            pyautogui.click(23,251)
            tm.sleep(1)

            os = tabela["DOCTO:"].values #Transforming dataframe to array
            preco = tabela["R$ DEVIDO:"].values
            cont = 0

            for nos in os:
                pyautogui.click(1382,537) #Localizar
                tm.sleep(0.5)
                pyautogui.typewrite(nos) #Digitar OS
                tm.sleep(0.5)
                pyautogui.press('enter') #Tabular
                pyautogui.press('enter') #Confirmar Localizar
                tm.sleep(0.5)
                pyautogui.press('tab') #Tabular para não baixar
                tm.sleep(0.5)
                pyautogui.press('enter') #Confirmar não biaxar
                tm.sleep(0.5)
                pyautogui.click(1381,598) #Clicar em Baixar
                tm.sleep(0.5)
                pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.press('enter')
                tm.sleep(0.5)
                pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.press('tab') #data
                tm.sleep(0.5)
                pyautogui.typewrite(dataCR)
                tm.sleep(0.5)
                ##pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.press('d')
                tm.sleep(0.5)
                pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.press('tab')
                tm.sleep(0.5)
                pyautogui.typewrite(tabela.iloc[cont,2])
                tm.sleep(0.5)
                pyautogui.press('enter')
                tm.sleep(0.5)
                pyautogui.click(953,590)
                tm.sleep(1)

                print(nos)
                print(tabela.iloc[cont,2])

                cont = cont + 1 #Line skip

        else:
            print("errado")
            runCR.configure(text="Insira a data novamente. (Data Inválida)", text_color="red")

    def destroy_cr():
        principal.deiconify()
        windowCR.destroy()

    titleCR = customtkinter.CTkLabel(windowCR, text="Baixar Contas a Receber", font=("Berlin Sans FB Demi", 24))
    titleCR.pack(pady=10)
    labelCR = customtkinter.CTkLabel(windowCR, text="Insira a data de pagamento das contas:", font=("Helvetica", 14))
    labelCR.pack(pady=5)
    entryCR = customtkinter.CTkEntry(windowCR, placeholder_text="                         dia/mes/ano", height=30, width=300, font=("Helvetica", 14), corner_radius=40, text_color="white", state="normal")
    entryCR.pack(pady=0)
    runCR = customtkinter.CTkLabel(windowCR, text="", font=("Helvetica", 10), text_color="green")
    runCR.pack(pady=1)

    startCR = customtkinter.CTkButton(windowCR, command=iniciar_processo_cr, text="Iniciar Processo", width=300, height=40, font=("Berlin Sans FB Demi", 18), corner_radius=40)
    startCR.pack(pady=10)
    destroyCR = customtkinter.CTkButton(windowCR, command=destroy_cr, text="< Voltar", font=("Helvetica", 10, "italic"), fg_color="#242424", text_color="white")
    destroyCR.place(anchor="center", x=200, y=220)

def TBWindow():
    windowTB = customtkinter.CTkToplevel(principal)
    principal.iconify()
    windowTB.title("Tabela Bergerson")
    windowTB.geometry("400x300")
    windowTB.resizable(False, False)

    def destroy_tb():
        principal.deiconify()
        windowTB.destroy()
    
    def iniciar_processo_tb():
        datai = entryTB1.get()
        dataf = entryTB2.get()
        # Abrindo arquivo DBF e criando o DataFrame
        path = ''r'C:\Users\Oficina\Desktop\Gustavo\Repositórios\Script-CIAF\Conversão\ORDEM-SERV.DBF'
        dbf = DBF(path)
        dataResult = pd.DataFrame(iter(dbf))

        #Filtro de colunas e transformando em datetime as datas
        tabela = dataResult[['NR_ORDEM','NRSERIE', 'DATAFECHA', 'VL_TOTAL', 'CODCLI', 'NOMECLIENT', 'STATUS']]
        tabela.loc[:, 'DATAFECHA'] = pd.to_datetime(tabela['DATAFECHA'], errors='coerce')

        #Inputs dos intervalos de data para aplicar filtro
        data_inicial = pd.to_datetime(datai)
        data_final = pd.to_datetime(dataf)

        #Filtros de cliente e data
        df_filtrado = tabela[(tabela['DATAFECHA'] >= data_inicial) & (tabela['DATAFECHA'] <= data_final)]
        df = df_filtrado[df_filtrado['CODCLI'].str.contains('147')]
        df = df[df['STATUS'].str.contains('FECHADA')]

        #Correção de dados e soma dos valores
        df['DATAFECHA'] = df['DATAFECHA'].fillna(pd.Timestamp('2024-01-01'))
        df['DATAFECHA'] = pd.to_datetime(df['DATAFECHA'])
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

        option = optionTB.get()

        if option == "Excel":
            df.to_excel(r"C:\Users\Oficina\Desktop\Extratos Bergerson - GScript\Tabela Bergerson.xlsx", index=False)
            statusTB.configure(text="Tabela Excel gerada com sucesso.")
        
        elif option == "PDF":
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
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOX', (0,0), (-1,-1), 1, colors.black),
                ('GRID', (0,0), (-1,-1), 0.25, colors.grey)
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
            doc = SimpleDocTemplate(r"C:\Users\Oficina\Desktop\Extratos Bergerson - GScript\Tabela Bergerson.pdf", pagesize=letter)

            elements = [title, date_range, Spacer(1,35) , table, Spacer(1,15),text_paragraph, Spacer(1,30), footer]
            doc.build(elements)
            statusTB.configure(text="Tabela PDF gerada com sucesso.")


    titleTB = customtkinter.CTkLabel(windowTB, text="Gerar tabela Bergerson", font=("Berlin Sans FB Demi", 24))
    titleTB.pack(pady=10)
    labelCR = customtkinter.CTkLabel(windowTB, text="Insira as datas de intervalo da consulta:", font=("Helvetica", 14))
    labelCR.pack(pady=5)
    entryTB1 = customtkinter.CTkEntry(windowTB, placeholder_text="      data inicial", height=30, width=145, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entryTB1.place(anchor="center", x=122, y=105)
    entryTB2 = customtkinter.CTkEntry(windowTB, placeholder_text="       data final", height=30, width=145, font=("Helvetica", 14, "italic"), corner_radius=40, text_color="white", state="normal")
    entryTB2.place(anchor="center", x=278, y=105)
    options = ["PDF", "Excel"]
    optionTB = customtkinter.CTkOptionMenu(windowTB, values=options, height=30, width=300, corner_radius=40, anchor="center")
    optionTB.place(anchor="center", x=200, y=155)
    startTB = customtkinter.CTkButton(windowTB, command=iniciar_processo_tb, text="Gerar Tabela", width=300, height=40, font=("Berlin Sans FB Demi", 18), corner_radius=40)
    startTB.place(anchor="center", x=200, y=210)
    statusTB = customtkinter.CTkLabel(windowTB, text="",font=("Helvetica", 10), text_color="green")
    statusTB.place(anchor="center", x=200 ,y=250)



    destroyCR = customtkinter.CTkButton(windowTB, command=destroy_tb, text="< Voltar", font=("Helvetica", 10, "italic"), width=80, height=30, fg_color="#242424", text_color="white", corner_radius=40)
    destroyCR.place(anchor="center", x=200, y=275)

def ORWindow():
    windowOR = customtkinter.CTkToplevel(principal)
    windowOR.geometry("1725x720+95+150")
    principal.iconify()
    windowOR.title("Orçamento")
    windowOR.resizable(False, False)

    global clientes
    global nomes_clientes  
    global telefone_clientes  
    global cpf_clientes
    global orcamentos
    global exec
    global pdf_frame
    global orcamento_atual
    orcamentos = []
    linhas = []

    today = datetime.now()
    data_atual = today.strftime("%d/%m/%Y")
    validade = today+timedelta(7)
    data_validade = validade.strftime("%d/%m/%Y")

    with open("orcamentos.csv", "r") as arquivo_orc:
        reader_orc = csv.DictReader(arquivo_orc)
        for row in reader_orc:
            orcamentos.append(row)

    with open("clientes.csv", "r") as arquivo_cli:
        reader = csv.DictReader(arquivo_cli)
        for row in reader:
            clientes.append(row)

    nomes_clientes = [cliente['nome'] for cliente in clientes]
    telefone_clientes = [cliente['telefone'] for cliente in clientes]
    cpf_clientes = [cliente['cpf'] for cliente in clientes]

    with open("config.txt", "r") as config:
        linhas = config.read().splitlines()
        orcamento = linhas[2]
        orcamento_atual = int(linhas[2])

    def tabview_pdf():
        global exec
        global orcamento_atual

        n_orc = int(orcamento_atual)
        df_orc = pd.read_csv("orcamentos.csv", encoding="ISO-8859-1")
        df_orc = df_orc.fillna(0)
        cliente = df_orc.loc[n_orc,"nome_cli"]
        file_name = "test/Orçamento "+str(n_orc)+" - "+cliente+".pdf"

        if not exec:
            global pdf_frame
            exportar_orcamento()

            pdf_frame = CTkPDFViewer(pdf_tab, file=file_name, page_width=425, page_height=500)
            pdf_frame.pack(fill="both", expand=True)
            exec = True
        else:
            exportar_orcamento()
            pdf_frame.configure(file=file_name)

    def novo_orcamento():
        global orcamento_atual

        new_btn.configure(state="disabled")

        with open("config.txt", "r") as config:
            linhas = config.read().splitlines()

        orcamento_atual = int(linhas[2])
        orcamento = int(linhas[2])
        orcamento += 1
        orcamento_atual += 1
        linhas[2] = str(orcamento)


        with open("config.txt", "w") as config:
            for linha in linhas:
                config.write(linha)
                config.write('\n')

        n_or.configure(text=orcamento)
        
        nome_cliente.set("Nome do Cliente")
        nome_cliente.bind("<FocusIn>", lambda e: nome_cliente.set(""))
        entry_telefone.delete("0", "end")
        entry_cpf.delete("0", "end")
        description_textbox.delete("0.0", "end")
        entry_prototp.delete("0", "end")
        entry_desenho.delete("0", "end")
        entry_molde.delete("0", "end")
        fundicao_entry.delete("0", "end")
        montagem_entry.delete("0", "end")
        acabamentos_entry.delete("0", "end")
        polimento_entry.delete("0", "end")
        limpeza_entry.delete("0", "end")
        cravacao_entry.delete("0", "end")
        ouro1k_entry.delete("0", "end")
        ouro750_entry.delete("0", "end")
        ourobranco_entry.delete("0", "end")
        prata_entry.delete("0", "end")
        rodio_entry.delete("0", "end")
        pedras_entry.delete("0", "end")
        servicost_entry.delete("0", "end")

    def exportar_orcamento():

        global valor_total
        global orcamento_atual
        
        def adicionar_servico(df_orc, df_horas, orcamento_atual, servico):
            if df_orc.loc[orcamento_atual, servico] != 0:
                df_horas.loc[len(df_horas),"Serviço"] = servico.capitalize()
                if df_orc.loc[orcamento_atual,"time_format"] == "Minutos":
                    df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"] = format((df_orc.loc[orcamento_atual,servico])/60, ".2f")
                    df_horas.loc[len(df_horas)-1,"Valor"] = format(float(df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"])*df_orc.loc[orcamento_atual,"preco_hora"], ".2f")
                else:
                    df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"] = (df_orc.loc[orcamento_atual,servico])
                    df_horas.loc[len(df_horas)-1,"Valor"] = format((df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"])*df_orc.loc[orcamento_atual,"preco_hora"], ".2f")

        with open(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\config.txt", "r") as config:
            linhas = config.read().splitlines()
            n_orc = int(linhas[2])

        df_orc = pd.read_csv(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\orcamentos.csv", encoding="ISO-8859-1")
        df_orc = df_orc.fillna(0)

        servicos = ["prototipagem", "desenho", "molde", "fundicao", "montagem", "acabamentos", "polimento", "limpeza", "cravacao"]

        df_horas = pd.DataFrame(columns=["Serviço", "Horas Trabalhadas", "Valor"])

        for servico in servicos:
            adicionar_servico(df_orc, df_horas, orcamento_atual, servico)

        total = df_horas["Valor"].astype(float).sum()
        total_horas = df_horas["Horas Trabalhadas"].replace('', '0').astype(float).sum()
        df_horas.loc[len(df_horas)] = ['Total', format(total_horas, ".2f"), format(total, ".2f")]

        df_cli = pd.read_csv(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\clientes.csv", encoding="ISO-8859-1")
        df_cli = df_cli.fillna(0)

        cliente = df_orc.loc[orcamento_atual,"nome_cli"]
        telefone = df_cli.loc[df_cli.index[df_cli["nome"]==cliente].tolist(), "telefone"].values[0]
        cpf = df_cli.loc[df_cli.index[df_cli["nome"]==cliente].tolist(), "cpf"].values[0]

        df_informacoes_cliente = pd.DataFrame({"Nome": [cliente], "Telefone": [telefone], "CPF/CNPJ": [cpf]})

        def adicionar_material(df_orc, df_materiais, orcamento_atual, material):
            if df_orc.loc[orcamento_atual, material] != 0:
                df_materiais.loc[len(df_materiais),"Material"] = material.capitalize()

                if material in ["pedras", "rodio", "servicos_terceiros"]:
                    df_materiais.loc[len(df_materiais)-1,"Valor"] = format(df_orc.loc[orcamento_atual, material], ".2f")
                else:
                    df_materiais.loc[len(df_materiais)-1,"Peso Utilizado"] = df_orc.loc[orcamento_atual,material]
                    df_materiais.loc[len(df_materiais)-1,"Valor"] = format(df_materiais.loc[len(df_materiais)-1,"Peso Utilizado"]*df_orc.loc[orcamento_atual,"cotacao"], ".2f")

        materiais = ["ouro1k", "ouro750", "ouro_branco", "prata", "pedras", "rodio", "servicos_terceiros"]

        df_materiais = pd.DataFrame(columns=["Material", "Peso Utilizado", "Valor"])

        for material in materiais:
            adicionar_material(df_orc, df_materiais, orcamento_atual, material)

        total = df_materiais["Valor"].astype(float).sum()
        df_materiais["Peso Utilizado"] = df_materiais["Peso Utilizado"].fillna("")
        df_materiais.loc[len(df_materiais)] = ["Total", "", format(total, ".2f")]

        df_informacoes_data = pd.DataFrame({"Data de Emissão": [df_orc.loc[orcamento_atual,"data_emissao"]], "Data de Validade": [df_orc.loc[orcamento_atual,"data_validade"]], "Validade": ["7 dias úteis"]})

        valor_subtotal = format(float(df_materiais.loc[len(df_materiais)-1, "Valor"]) + float(df_horas.loc[len(df_horas)-1, "Valor"]))
        valor_lucro = format(float(df_horas.loc[len(df_horas)-1, "Valor"]) * float(df_orc.loc[orcamento_atual, "taxa_lucro"]), ".2f")
        valor_frete = format(float(df_orc.loc[orcamento_atual, "frete"]), ".2f")
        valor_desconto = format(float(df_orc.loc[orcamento_atual, "desconto"]), ".2f")
        valor_total = format(float(df_materiais.loc[len(df_materiais)-1, "Valor"]) + float(df_horas.loc[len(df_horas)-1, "Valor"]) + float(valor_lucro) + float(valor_frete) - float(valor_desconto), ".2f")
        df_valorfinal = pd.DataFrame({"Preços":["Subtotal", "Lucro:  "+str(format(df_orc.loc[orcamento_atual, "taxa_lucro"] * 100, ".1f"))+"%", "Frete","Desconto", "Total"], "R$": [valor_subtotal, valor_lucro, valor_frete, valor_desconto, valor_total]})

        df_descricao = pd.DataFrame({"Descrição do Projeto": [df_orc.loc[orcamento_atual,"descricao"]]})

        # Criação do PDF com reportlab
        styles = getSampleStyleSheet()

        # Criação dos diferentes estilos utilizados nos documentos
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=20,
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

        nametable_style = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ])

        table_style = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey)
        ])

        table_style_final = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONT', (-1, -1), (-1, -1), 'Helvetica-Bold', 12),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey)
        ])


        def myFirstPage(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            canvas.drawString(inch, 0.75 * inch, "OR-"+str(orcamento_atual))
            canvas.drawString(7*inch, 0.75 * inch, "RDA Design")
            canvas.drawString(6.79*inch, 0.9 * inch, "_________________")
            canvas.restoreState()

        caminho = r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\test\Orçamento " + str(orcamento_atual) + " - " + cliente + ".pdf"
        doc = BaseDocTemplate(caminho, pagesize=letter)
        frame = Frame(inch, inch, 6.5 * inch, 9.7 * inch, id='normal')
        template = PageTemplate(id='test', frames=frame, onPage=myFirstPage)
        doc.addPageTemplates([template])

        title_text = "ORÇAMENTO Nº" + str(orcamento_atual)
        title = Paragraph(title_text, title_style)

        linha = Drawing(2000, 10)  # Largura e altura da linha
        line = Line(-42, 0, 497, 0, strokeColor=colors.black, strokeWidth=1)  # Coordenadas, cor e espessura
        linha.add(line)

        data_clientes = [df_informacoes_cliente.columns.to_list()] + df_informacoes_cliente.values.tolist()
        tabela_clientes = Table(data_clientes)
        tabela_clientes.setStyle(nametable_style)
        tabela_clientes._argW[0] = 2.5 * inch
        tabela_clientes._argW[1] = 2.5 * inch
        tabela_clientes._argW[2] = 2.5 * inch

        data_datas = [df_informacoes_data.columns.to_list()] + df_informacoes_data.values.tolist()
        tabela_datas = Table(data_datas)
        tabela_datas.setStyle(nametable_style)
        tabela_datas._argW[0] = 2.5 * inch
        tabela_datas._argW[1] = 2.5 * inch
        tabela_datas._argW[2] = 2.5 * inch

        data_descricao = [df_descricao.columns.to_list()] + df_descricao.values.tolist()
        tabela_descricao = Table(data_descricao)
        tabela_descricao.setStyle(table_style)
        tabela_descricao._argW[0] = 7.5 * inch

        data_servicos = [df_horas.columns.to_list()] + df_horas.values.tolist()
        tabela_servicos = Table(data_servicos)
        tabela_servicos.setStyle(table_style)
        tabela_servicos._argW[0] = 2.5 * inch
        tabela_servicos._argW[1] = 2.5 * inch
        tabela_servicos._argW[2] = 2.5 * inch

        data_materiais = [df_materiais.columns.to_list()] + df_materiais.values.tolist()
        tabela_materiais = Table(data_materiais)
        tabela_materiais.setStyle(table_style)
        tabela_materiais._argW[0] = 2.5 * inch
        tabela_materiais._argW[1] = 2.5 * inch
        tabela_materiais._argW[2] = 2.5 * inch

        data_final = [df_valorfinal.columns.to_list()] + df_valorfinal.values.tolist()
        tabela_valorfinal = Table(data_final)
        tabela_valorfinal.setStyle(table_style_final)
        tabela_valorfinal._argW[0] = 5 * inch
        tabela_valorfinal._argW[1] = 2.5 * inch

        elements = [title, Spacer(1, 10), tabela_descricao, Spacer(1, 5), linha, tabela_clientes, tabela_datas, linha, Spacer(1, 15),
                    tabela_servicos, Spacer(1, 20), tabela_materiais, Spacer(20, 20), tabela_valorfinal]
        doc.build(elements)
        
    def salvar_orcamento():
        try:
            with open("config.txt", "r") as config:
                linhas = config.read().splitlines()

            if orcamento_atual < len(orcamentos):
                n_orc = orcamento_atual
            else:
                n_orc = int(linhas[2])

            data_emissao = entry_dataemissao.get()
            data_validade = entry_datavalidade.get()
            nome_cli = nome_cliente.get()
            descricao = description_textbox.get("1.0", "end")
            time_format = tempo.get()
            prototipagem = entry_prototp.get()
            desenho = entry_desenho.get()
            molde = entry_molde.get()
            fundicao = fundicao_entry.get()
            montagem = montagem_entry.get()
            acabamentos = acabamentos_entry.get()
            polimento = polimento_entry.get()
            limpeza = limpeza_entry.get()
            cravacao = cravacao_entry.get()
            ouro1k = ouro1k_entry.get()
            ouro750 =  ouro750_entry.get()
            ouro_branco = ourobranco_entry.get()
            pedras = pedras_entry.get()
            prata = prata_entry.get()
            rodio = rodio_entry.get()
            servicos_terceiros = servicost_entry.get()
            cotacao = cotacao_entry.get()
            preco_hora = precohora_entry.get()
            taxa_lucro = lucro_entry.get()
            taxa_lucro = float(taxa_lucro) / 100
            frete = entry_frete.get()
            desconto = entry_desconto.get()

            orcamento = {"n_orc": n_orc,
                        "data_emissao": data_emissao,
                        "data_validade": data_validade,
                        "nome_cli":nome_cli,
                        "descricao":descricao,
                        "time_format": time_format,
                        "prototipagem":prototipagem,
                        "desenho":desenho,
                        "molde":molde,
                        "fundicao":fundicao,
                        "montagem":montagem,
                        "acabamentos":acabamentos,
                        "polimento":polimento,
                        "limpeza":limpeza,
                        "cravacao":cravacao,
                        "ouro1k":ouro1k,
                        "ouro750":ouro750,
                        "ouro_branco":ouro_branco,
                        "pedras":pedras,
                        "prata":prata,
                        "rodio":rodio,
                        "servicos_terceiros":servicos_terceiros,
                        "cotacao":cotacao,
                        "preco_hora":preco_hora,
                        "taxa_lucro":taxa_lucro,
                        "frete": frete,
                        "desconto": desconto
                        }
            
            if orcamento_atual < len(orcamentos):
                orcamentos[orcamento_atual] = orcamento
            else:
                orcamentos.append(orcamento)

            with open("orcamentos.csv", mode='w', newline='') as orc:
                writer = csv.DictWriter(orc, fieldnames=["n_orc","data_emissao","data_validade","nome_cli","descricao","time_format","prototipagem","desenho","molde","fundicao","montagem","acabamentos","polimento","limpeza","cravacao","ouro1k","ouro750","ouro_branco","pedras","prata","rodio","servicos_terceiros","cotacao","preco_hora","taxa_lucro","frete","desconto"])
                writer.writeheader()
                for orcamento in orcamentos:
                    writer.writerow(orcamento)
            
            new_btn.configure(state="normal")

            print("Orçamento Salvo")
        except Exception as e:
            print("Erro em Salvamento")
        
    def search_cliente(event):
        value = event.widget.get()
        if value == '':
            nome_cliente['values'] = nomes_clientes
        else:
            filtro = []
            for item in nomes_clientes:
                if value.lower() in item.lower():
                    filtro.append(item)
            nome_cliente.configure(values=filtro)

    def adicionar_cliente_cmd():
        global clientes
        global nomes_clientes  
        global telefone_clientes  
        global cpf_clientes  
        nome = nome_cliente.get()
        telefone = entry_telefone.get()
        cpf = entry_cpf.get()
        adicionar_cliente(nome, telefone, cpf)
        nomes_clientes = [cliente['nome'] for cliente in clientes]  
        telefone_clientes = [cliente['telefone'] for cliente in clientes]  
        cpf_clientes = [cliente['cpf'] for cliente in clientes]  
        nome_cliente.configure(values=nomes_clientes)

    def adicionar_cliente(nome, telefone, cpf):
        cliente = {"nome": nome, "telefone": telefone, "cpf": cpf}
        clientes.append(cliente)

        with open("clientes.csv", mode='w', newline='') as arquivo:
            writer = csv.DictWriter(arquivo, fieldnames=["nome", "telefone", "cpf"])
            writer.writeheader()
            for cliente in clientes:
                writer.writerow(cliente)

    def preencher_campos(event):
        idx = nomes_clientes.index(nome_cliente.get())
        telefone = telefone_clientes[idx]
        cpf = cpf_clientes[idx]
        entry_telefone.delete(0, 'end')
        entry_telefone.insert(0, telefone)
        entry_cpf.delete(0, 'end')
        entry_cpf.insert(0, cpf)

    def switch_tempo():
        medida = tempo.get()
        if medida == "Horas":
            tempo.configure(text="Unidade de Tempo: Horas")
            entry_prototp.configure(placeholder_text="Horas")
            entry_desenho.configure(placeholder_text="Horas")
            entry_molde.configure(placeholder_text="Horas")
            fundicao_entry.configure(placeholder_text="Horas")
            montagem_entry.configure(placeholder_text="Horas")
            acabamentos_entry.configure(placeholder_text="Horas")
            polimento_entry.configure(placeholder_text="Horas")
            limpeza_entry.configure(placeholder_text="Horas")
            cravacao_entry.configure(placeholder_text="Horas")
        else:
            tempo.configure(text="Unidade de Tempo: Minutos")
            entry_prototp.configure(placeholder_text="Minutos")
            entry_desenho.configure(placeholder_text="Minutos")
            entry_molde.configure(placeholder_text="Minutos")
            fundicao_entry.configure(placeholder_text="Minutos")
            montagem_entry.configure(placeholder_text="Minutos")
            acabamentos_entry.configure(placeholder_text="Minutos")
            polimento_entry.configure(placeholder_text="Minutos")
            limpeza_entry.configure(placeholder_text="Minutos")
            cravacao_entry.configure(placeholder_text="Minutos")

    def sliding_cotacao(value):
        value = format(value, '.2f')
        cotacao_entry.delete("0", 'end')
        cotacao_entry.insert("0", value)
    
    def entry_cotacao(event):
        cotacao = cotacao_entry.get()
        cotacao_slider.set(float(cotacao))

    def sliding_hora(value):
        value = format(value, '.2f')
        precohora_entry.delete("0", 'end')
        precohora_entry.insert("0", value)
        
    def entry_hora(event):
        cotacao = precohora_entry.get()
        precohora_slider.set(float(cotacao))
    
    def sliding_lucro(value): #revisar

        value = format(value, '.2f')
        lucro_entry.delete("0", 'end')
        lucro_entry.insert("0", value)

    def entry_lucro(event):
        cotacao = lucro_entry.get()
        lucro_slider.set(float(cotacao))

    def search():
        global orcamento_atual
        searching = search_entry.get()
        orcamento_atual = int(searching)
        
        try:
            entry_dataemissao.insert(0, (orcamentos[orcamento_atual]['data_emissao']))
            entry_dataemissao.delete(0,'end')
            entry_dataemissao.insert(0, (orcamentos[orcamento_atual]['data_emissao']))
            entry_datavalidade.delete(0,'end')
            entry_datavalidade.insert(0, (orcamentos[orcamento_atual]['data_validade']))
            nome_cliente.set(orcamentos[orcamento_atual]['nome_cli'])
            preencher_campos(orcamentos[orcamento_atual]['nome_cli'])
            description_textbox.delete('0.0','end')
            description_textbox.insert('0.0', (orcamentos[orcamento_atual]['descricao']))
            if (orcamentos[orcamento_atual]['time_format']) == "Minutos":
                tempo.deselect()
            else:
                tempo.select()
            entry_prototp.delete(0,'end')
            entry_prototp.insert(0, (orcamentos[orcamento_atual]['prototipagem']))
            entry_desenho.delete(0,'end')
            entry_desenho.insert(0, (orcamentos[orcamento_atual]['desenho']))
            entry_molde.delete(0,'end')
            entry_molde.insert(0, (orcamentos[orcamento_atual]['molde']))
            fundicao_entry.delete(0,'end')
            fundicao_entry.insert(0, (orcamentos[orcamento_atual]['fundicao']))
            montagem_entry.delete(0,'end')
            montagem_entry.insert(0, (orcamentos[orcamento_atual]['montagem']))
            acabamentos_entry.delete(0,'end')
            acabamentos_entry.insert(0, (orcamentos[orcamento_atual]['acabamentos']))
            polimento_entry.delete(0,'end')
            polimento_entry.insert(0, (orcamentos[orcamento_atual]['polimento']))
            limpeza_entry.delete(0,'end')
            limpeza_entry.insert(0, (orcamentos[orcamento_atual]['limpeza']))
            cravacao_entry.delete(0,'end')
            cravacao_entry.insert(0, (orcamentos[orcamento_atual]['cravacao']))
            ouro1k_entry.delete(0,'end')
            ouro1k_entry.insert(0, (orcamentos[orcamento_atual]['ouro1k']))
            ouro750_entry.delete(0,'end')
            ouro750_entry.insert(0, (orcamentos[orcamento_atual]['ouro750']))
            ourobranco_entry.delete(0,'end')
            ourobranco_entry.insert(0, (orcamentos[orcamento_atual]['ouro_branco']))
            pedras_entry.delete(0,'end')
            pedras_entry.insert(0, (orcamentos[orcamento_atual]['pedras']))
            prata_entry.delete(0,'end')
            prata_entry.insert(0, (orcamentos[orcamento_atual]['prata']))
            rodio_entry.delete(0,'end')
            rodio_entry.insert(0, (orcamentos[orcamento_atual]['rodio']))
            servicost_entry.delete(0,'end')
            servicost_entry.insert(0, (orcamentos[orcamento_atual]['servicos_terceiros']))
            cotacao_entry.delete(0,'end')
            cotacao_entry.insert(0, (orcamentos[orcamento_atual]['cotacao']))
            precohora_entry.delete(0,'end')
            precohora_entry.insert(0, (orcamentos[orcamento_atual]['preco_hora']))
            n_or.configure(text=orcamento_atual)

            print("Busca realizada com sucesso")
        except:
            def icon():
                aviso.destroy()
                search_entry.delete(0, 'end')
            aviso = customtkinter.CTkToplevel(windowOR)
            aviso.geometry("300x150")
            aviso.deiconify()
            aviso_label = customtkinter.CTkLabel(aviso, text="O arquivo pesquisado não existe.")
            aviso_label.pack(pady=30)
            aviso_btn = customtkinter.CTkButton(aviso, text="Tentar novamente", command=icon, corner_radius=40)
            aviso_btn.pack(pady=5)
            tm.sleep(1)
            
    def scrolldown_t(event):
        tempo_inframe._parent_canvas.yview_moveto(1)
    
    def scrolldown_o(event):
        material_inframe._parent_canvas.yview_moveto(0.35)

    def scrolldown_r(event):
        material_inframe._parent_canvas.yview_moveto(1)

    def tabulation_d(event):
        entry_prototp.focus_force()
        return "break"

    def destroy_or():
        principal.deiconify()
        windowOR.destroy()

    top_frame = customtkinter.CTkFrame(master=windowOR, width=375, height=150, corner_radius=40)
    top_frame.place(y=30, x=30, anchor="nw")
    title_label = customtkinter.CTkButton(top_frame, text="Orçamento",width=200, height=50, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 22), hover=False).place(y=40, x=120, anchor="center")
    n_or = customtkinter.CTkButton(top_frame, text=orcamento,width=80, height=50, corner_radius=40, fg_color="#1f6aa5",font=("Arial", 20, "bold"), hover=False)
    n_or.place(y=40, x=306, anchor="center")
    labelemissao = customtkinter.CTkLabel(top_frame, text="Data da emissão:", font=("Helvetica", 12))
    labelemissao.place(y=90, x=100, anchor="center")
    entry_dataemissao = customtkinter.CTkEntry(top_frame,justify="center", placeholder_text="  /  /  ", height=30, width=150, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_dataemissao.place(y=120, x=100, anchor="center")
    entry_dataemissao.insert(0, data_atual)
    labelvalidade = customtkinter.CTkLabel(top_frame, text="Válido até:", font=("Helvetica", 12))
    labelvalidade.place(y=90, x=280, anchor="center")
    entry_datavalidade = customtkinter.CTkEntry(top_frame,justify="center", placeholder_text="  /  /  ", height=30, width=150, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_datavalidade.place(y=120, x=280, anchor="center")
    entry_datavalidade.insert(0, data_validade)

    h_frame = customtkinter.CTkFrame(master=windowOR, width=375, height=190, corner_radius=40)
    h_frame.place(x=30,y=190,anchor="nw")
    nome_cliente = customtkinter.CTkComboBox(h_frame, command=preencher_campos, values=nomes_clientes ,width=335, height=40, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), button_color="#1f6aa5", justify="center")
    nome_cliente.place(y=35, x=187, anchor="center")
    nome_cliente.set("Nome do Cliente")
    nome_cliente.bind("<FocusIn>", lambda e: nome_cliente.set(""))
    nome_cliente.bind("<KeyRelease>",search_cliente)
    labelTelefone = customtkinter.CTkLabel(h_frame, text="Telefone:", font=("Helvetica", 12))
    labelTelefone.place(y=80, x=100, anchor="center")
    entry_telefone = customtkinter.CTkEntry(h_frame, placeholder_text="(  )            -    ", height=30, width=150, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_telefone.place(y=110, x=100, anchor="center")
    labelcpf = customtkinter.CTkLabel(h_frame, text="CPF/CNPJ", font=("Helvetica", 12))
    labelcpf.place(y=80, x=280, anchor="center")
    entry_cpf = customtkinter.CTkEntry(h_frame, placeholder_text="       .       .      -", height=30, width=150, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_cpf.place(y=110, x=280, anchor="center")
    addcliente = customtkinter.CTkButton(h_frame,text="Adicionar Cliente", command=adicionar_cliente_cmd, width=335, height=30, corner_radius=40, font=("Helvetica", 14,"bold")).place(x=187,y=155,anchor="center")

    description_frame = customtkinter.CTkFrame(master=windowOR, width=375, height=250, corner_radius=40)
    description_frame.place(anchor="nw", x=30, y=390)
    description_label= customtkinter.CTkButton(description_frame, text="Descrição do Orçamento",width=334, height=40, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), hover=False)
    description_label.place(y=35, x=187, anchor="center")
    description_textbox = customtkinter.CTkTextbox(description_frame, width=334, height=160, corner_radius=20, fg_color="#242424", border_spacing=0, border_color="#1f6aa5", border_width=1)
    description_textbox.place(x=187, y=152, anchor="center")

    hours_frame = customtkinter.CTkFrame(master=windowOR, width=375, height= 610, corner_radius=40)
    hours_frame.place(anchor="nw", x=415, y=30)
    hourswrk_label = customtkinter.CTkButton(hours_frame, text="Tempo Trabalhado",width=334, height=40, fg_color="#242424", border_color="#397445", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), hover=False)
    hourswrk_label.place(y=35, x=187, anchor="center")
    tempo_inframe = customtkinter.CTkScrollableFrame(master=hours_frame, fg_color="#242424", width=295, height=485, corner_radius=20,border_color="#397445", border_width=1)
    tempo_inframe.place(anchor="center", x=187, y=330)
    switch_frame = customtkinter.CTkFrame(master=tempo_inframe, width=275, height=40, corner_radius=10, fg_color="#397445")
    switch_frame.pack(pady=(10,5))
    tempo_var = customtkinter.StringVar(value="Minutos")
    tempo = customtkinter.CTkSwitch(master=switch_frame, width=260, text="Unidade de Tempo: Minutos", command=switch_tempo, onvalue="Horas", offvalue="Minutos", progress_color="#242424", variable=tempo_var, font=("Helvetica", 14, "bold"))
    tempo.pack(pady=5, padx=(11,0))
    prototp_label = customtkinter.CTkLabel(tempo_inframe, text="Prototipagem:", font=("Helvetica",14,"bold"))
    prototp_label.pack(pady=(10,5), padx=(3,5))
    entry_prototp = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    entry_prototp.pack(pady=(5,10), padx=(3,5))
    desenho_label = customtkinter.CTkLabel(tempo_inframe, text="Desenho:", font=("Helvetica",14,"bold"))
    desenho_label.pack(pady=(10,5), padx=(3,5))
    entry_desenho = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    entry_desenho.pack(pady=(5,10), padx=(3,5))
    molde_label = customtkinter.CTkLabel(tempo_inframe, text="Molde em Cera:", font=("Helvetica",14,"bold"))
    molde_label.pack(pady=(10,5), padx=(3,5))
    entry_molde = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    entry_molde.pack(pady=(5,10), padx=(3,5))
    fundicao_label = customtkinter.CTkLabel(tempo_inframe, text="Fundição:", font=("Helvetica",14,"bold"))
    fundicao_label.pack(pady=(10,5), padx=(3,5))
    fundicao_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    fundicao_entry.pack(pady=(5,10), padx=(3,5))
    montagem_label = customtkinter.CTkLabel(tempo_inframe, text="Montagem:", font=("Helvetica",14,"bold"))
    montagem_label.pack(pady=(10,5), padx=(3,5))
    montagem_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    montagem_entry.pack(pady=(5,10), padx=(3,5))
    acabamentos_label = customtkinter.CTkLabel(tempo_inframe, text="Acabamentos:", font=("Helvetica",14,"bold"))
    acabamentos_label.pack(pady=(10,5), padx=(3,5))
    acabamentos_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    acabamentos_entry.pack(pady=(5,10), padx=(3,5))
    polimento_label = customtkinter.CTkLabel(tempo_inframe, text="Polimento:", font=("Helvetica",14,"bold"))
    polimento_label.pack(pady=(10,5), padx=(3,5))
    polimento_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    polimento_entry.pack(pady=(5,10), padx=(3,5))
    limpeza_label = customtkinter.CTkLabel(tempo_inframe, text="Limpeza:", font=("Helvetica",14,"bold"))
    limpeza_label.pack(pady=(10,5), padx=(3,5))
    limpeza_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    limpeza_entry.pack(pady=(5,10), padx=(3,5))
    cravacao_label = customtkinter.CTkLabel(tempo_inframe, text="Cravação:", font=("Helvetica",14,"bold"))
    cravacao_label.pack(pady=(10,5), padx=(3,5))
    cravacao_entry = customtkinter.CTkEntry(tempo_inframe, placeholder_text="minutos", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    cravacao_entry.pack(pady=(5,10), padx=(3,5))
    separator = customtkinter.CTkLabel(tempo_inframe, text="___________________________________________________", font=("Helvetica",8), text_color="#343638")
    separator.pack(anchor="center", pady=(3,8), padx=(3,5))
    precohora_label = customtkinter.CTkLabel(tempo_inframe, text="Preço da Hora Trabalhada:", font=("Helvetica",14,"bold"))
    precohora_label.pack(anchor="center", pady=10, padx=(3,5))
    precohora_frame = customtkinter.CTkFrame(master=tempo_inframe, width=300, height= 50, fg_color="#397445", corner_radius=40)
    precohora_frame.pack(pady=(5,10), padx=(3,5))
    precohora_slider = customtkinter.CTkSlider(precohora_frame, command=sliding_hora,width=182, height=20, from_=0, to=200, number_of_steps=2000, button_color="#d5d9de", button_hover_color="white")
    precohora_slider.place(anchor="w", x=10,y=25)
    precohora_entry = customtkinter.CTkEntry(precohora_frame, placeholder_text="R$", justify="center", height=30, width=80, font=("Helvetica", 12,"bold"), corner_radius=40, border_color="#2F5D39", border_width=1, text_color="white", state="normal")
    precohora_entry.place(anchor="w", x=197,y=25)
    precohora_entry.bind("<FocusOut>", entry_hora)

    material_frame = customtkinter.CTkFrame(master=windowOR, width=375, height= 400, corner_radius=40)
    material_frame.place(anchor="nw", x=800, y=30)
    material_label = customtkinter.CTkButton(material_frame, text="Material Utilizado",width=334, height=40, fg_color="#242424", border_color="#70B8B8", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), hover=False)
    material_label.place(y=35, x=187, anchor="center")
    material_inframe = customtkinter.CTkScrollableFrame(master=material_frame, fg_color="#242424", width=295, height=275, corner_radius=20,border_color="#70B8B8", border_width=1)
    material_inframe.place(anchor="center", x=187, y=225)
    cotacao_frame = customtkinter.CTkFrame(master=material_inframe, width=275, height=40, corner_radius=10, fg_color="#70B8B8")
    cotacao_frame.pack(pady=(10,5))
    cotacao = customtkinter.CTkLabel(cotacao_frame, width=260, text="Ouro:", font=("Helvetica", 14, "bold"))
    cotacao.pack(pady=5, padx=(11,0))
    ouro1k_label = customtkinter.CTkLabel(material_inframe, text="Ouro 1000:", font=("Helvetica",14,"bold"))
    ouro1k_label.pack(pady=(10,5), padx=(3,5))
    ouro1k_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="g", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    ouro1k_entry.pack(pady=(5,10), padx=(3,5))
    ouro750_label = customtkinter.CTkLabel(material_inframe, text="Ouro 750:", font=("Helvetica",14,"bold"))
    ouro750_label.pack(pady=(10,5), padx=(3,5))
    ouro750_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="g", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    ouro750_entry.pack(pady=(5,10), padx=(3,5))
    ourobranco_label = customtkinter.CTkLabel(material_inframe, text="Ouro Branco:", font=("Helvetica",14,"bold"))
    ourobranco_label.pack(pady=(10,5), padx=(3,5))
    ourobranco_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="g", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    ourobranco_entry.pack(pady=(5,10), padx=(3,5))
    prata_label = customtkinter.CTkLabel(material_inframe, text="Prata:", font=("Helvetica",14,"bold"))
    prata_label.pack(pady=(10,5), padx=(3,5))
    prata_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="g", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    prata_entry.pack(pady=(5,10), padx=(3,5))
    rodio_label = customtkinter.CTkLabel(material_inframe, text="Ródio:", font=("Helvetica",14,"bold"))
    rodio_label.pack(pady=(10,5), padx=(3,5))
    rodio_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="R$", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    rodio_entry.pack(pady=(5,10), padx=(3,5))
    pedras_label = customtkinter.CTkLabel(material_inframe, text="Pedras:", font=("Helvetica",14,"bold"))
    pedras_label.pack(pady=(10,5), padx=(3,5))
    pedras_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="R$", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    pedras_entry.pack(pady=(5,10), padx=(3,5))
    servicost_label = customtkinter.CTkLabel(material_inframe, text="Serviços de Terceiros:", font=("Helvetica",14,"bold"))
    servicost_label.pack(pady=(10,5), padx=(3,5))
    servicost_entry = customtkinter.CTkEntry(material_inframe, placeholder_text="R$", justify="center", height=30, width=275, font=("Helvetica", 14,"bold"), corner_radius=40, border_color="#565b5e", border_width=1, text_color="white", state="normal")
    servicost_entry.pack(pady=(5,10), padx=(3,5))

    precos_frame = customtkinter.CTkFrame(master=windowOR, width=375, height= 200, corner_radius=40)
    precos_frame.place(anchor="nw", x=800, y=440)
    precos_label = customtkinter.CTkButton(precos_frame, text="Preçificação",width=334, height=40, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), hover=False)
    precos_label.place(y=35, x=187, anchor="center")
    precos_inframe = customtkinter.CTkFrame(master=precos_frame, fg_color="#242424", width=334, height=120, corner_radius=20,border_color="#1f6aa5", border_width=1)
    precos_inframe.place(anchor="center", x=187, y=125)
    cotacaometal_label = customtkinter.CTkLabel(precos_inframe, text="Cotação do Metal utilizado:", font=("Helvetica",14,"bold"))
    cotacaometal_label.place(anchor="center", x=167, y=25)
    cotacaopreco_frame = customtkinter.CTkFrame(master=precos_inframe, width=300, height= 50, fg_color="#1f6aa5", corner_radius=40)
    cotacaopreco_frame.place(anchor="center", x=167, y=70)
    cotacao_slider = customtkinter.CTkSlider(cotacaopreco_frame, command=sliding_cotacao,width=195, height=20, from_=0, to=410, number_of_steps=4100, button_color="#d5d9de", button_hover_color="white")
    cotacao_slider.place(anchor="w", x=10,y=25)
    cotacao_entry = customtkinter.CTkEntry(cotacaopreco_frame, placeholder_text="R$", justify="center", height=30, width=80, font=("Helvetica", 12,"bold"), corner_radius=40, border_color="#144870", border_width=1, text_color="white", state="normal")
    cotacao_entry.place(anchor="w", x=210,y=25)
    cotacao_entry.bind("<FocusOut>", entry_cotacao)

    exec = False
    tabs = customtkinter.CTkTabview(windowOR, width=500, height=625, corner_radius=40, command=tabview_pdf)
    tabs.place(anchor="nw", x=1185, y=13)

    precos_tab = tabs.add("Preços")
    #tabelas_tab = tabs.add("Tabelas")
    pdf_tab = tabs.add("PDF")

    precos1_label = customtkinter.CTkButton(precos_tab, text="Preçificação",width=410, height=40, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 18), hover=False)
    precos1_label.place(anchor="nw", x=10, y=10)
    precostab_frame = customtkinter.CTkFrame(master=precos_tab, width=408, height=440, fg_color="#242424", corner_radius=20, border_color="#1f6aa5", border_width=1)
    precostab_frame.place(anchor="nw", x=10, y=60)
    cotacaometal_label = customtkinter.CTkLabel(precostab_frame, text="Lucro - Sobre a mão de obra.", font=("Helvetica",14,"bold"))
    cotacaometal_label.place(anchor="center", x=207, y=30)
    lucro_frame = customtkinter.CTkFrame(master=precostab_frame, width=300, height= 50, fg_color="#1f6aa5", corner_radius=40)
    lucro_frame.place(anchor="center", x=207, y=75)
    lucro_slider = customtkinter.CTkSlider(lucro_frame, command=sliding_lucro,width=195, height=20, from_=0, to=100, number_of_steps=200, button_color="#d5d9de", button_hover_color="white")
    lucro_slider.place(anchor="w", x=10,y=25)
    lucro_slider.set(12)
    lucro_entry = customtkinter.CTkEntry(lucro_frame, placeholder_text="%", justify="center", height=30, width=80, font=("Helvetica", 12,"bold"), corner_radius=40, border_color="#144870", border_width=1, text_color="white", state="normal")
    lucro_entry.place(anchor="w", x=210,y=25)
    lucro_entry.bind("<FocusOut>", entry_lucro)
    labelfrete = customtkinter.CTkLabel(precostab_frame, text="Frete", font=("Helvetica", 14, "bold"))
    labelfrete.place(y=130, x=130, anchor="center")
    entry_frete = customtkinter.CTkEntry(precostab_frame,justify="center", placeholder_text="R$", height=30, width=140, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_frete.place(y=145, x=60, anchor="nw")
    labeldesconto = customtkinter.CTkLabel(precostab_frame, text="Desconto", font=("Helvetica", 14, "bold"))
    labeldesconto.place(y=130, x=285, anchor="center")
    entry_desconto = customtkinter.CTkEntry(precostab_frame,justify="center", placeholder_text="R$", height=30, width=140, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entry_desconto.place(y=145, x=215, anchor="nw")
    separator_p = customtkinter.CTkLabel(precos_tab, text="_______________________________________________________________", font=("Helvetica",8), text_color="#343638", bg_color="#242424")
    separator_p.place(anchor="center", x=215, y=265)
    valortotal_label = customtkinter.CTkButton(precos_tab, text="R$ 1400,00",width=300, height=40, bg_color="#242424", fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=25, font=("Berlin Sans FB Demi", 22), hover=False)
    valortotal_label.place(anchor="center", x=217, y=315)

    new_img = customtkinter.CTkImage(light_image=Image.open("img/new.png"), dark_image=Image.open("img/new.png"), size=(17,17))
    save_img = customtkinter.CTkImage(light_image=Image.open("img/save.png"), dark_image=Image.open("img/save.png"), size=(17,17))
    export_img = customtkinter.CTkImage(light_image=Image.open("img/export.png"), dark_image=Image.open("img/export.png"), size=(17,17))
    print_img = customtkinter.CTkImage(light_image=Image.open("img/print.png"), dark_image=Image.open("img/print.png"), size=(17,17))
    search_img = customtkinter.CTkImage(light_image=Image.open("img/search.png"), dark_image=Image.open("img/search.png"), size=(15,15))


    new_btn = customtkinter.CTkButton(windowOR, text="Novo",width=250, height=40, command=novo_orcamento, font=("Berlin Sans FB Demi", 22), corner_radius=40, image=new_img)
    new_btn.place(anchor="center", x=155, y=680)
    save_btn = customtkinter.CTkButton(windowOR, text="Salvar",width=250, height=40, command=salvar_orcamento, font=("Berlin Sans FB Demi", 22), corner_radius=40, image=save_img)
    save_btn.place(anchor="center", x=415, y=680)
    export_btn = customtkinter.CTkButton(windowOR, text="Exportar",width=250, height=40, command=exportar_orcamento, font=("Berlin Sans FB Demi", 22), corner_radius=40, image=export_img)
    export_btn.place(anchor="center", x=675, y=680)
    print_btn = customtkinter.CTkButton(windowOR, text="Imprimir",width=250, height=40, font=("Berlin Sans FB Demi", 22), corner_radius=40, image=print_img)
    print_btn.place(anchor="center", x=935, y=680)
    #share_btn = customtkinter.CTkButton(windowOR, text="Exportar",width=250, height=40, font=("Berlin Sans FB Demi", 22), corner_radius=40)
    #share_btn.place(anchor="center", x=675, y=680)
    search_frame =customtkinter.CTkFrame(master=windowOR, width=250, height=40, corner_radius=40, fg_color="#242424", border_width=1, border_color="#1f6aa5")
    search_frame.place(anchor="center", x=1300, y=680)
    search_entry = customtkinter.CTkEntry(search_frame, placeholder_text="Pesquisar", width=180, height=30, fg_color="#242424", border_color="#242424")
    search_entry.place(anchor="center", x=105, y=20)
    search_btn = customtkinter.CTkButton(search_frame, width=20, height=30, command=search, text="", corner_radius=50, image=search_img)
    search_btn.place(anchor="center", x=218, y=20)

    destroyOR = customtkinter.CTkButton(windowOR, command=destroy_or, text="Voltar", width=250, height=40, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 20))
    destroyOR.place(anchor="center", x=1560, y=680)
    description_textbox.bind("<Tab>", tabulation_d)
    montagem_entry.bind("<FocusOut>", scrolldown_t)
    ouro750_entry.bind("<FocusOut>", scrolldown_o)
    rodio_entry.bind("<FocusOut>", scrolldown_r)
    entry_dataemissao.focus()

def settings():
    settings = customtkinter.CTkToplevel(principal)
    principal.iconify()
    settings.geometry("400x300")
    settings.resizable(False, False)
    settings.title("Configurações")

    with open("config.txt", "r") as config:
        linhas = config.readlines()
    
    def destroy_settings():
        principal.deiconify()
        settings.destroy()

    def fileselector_CR():
        linhas[0] = filedialog.askdirectory()
        with open("config.txt", "w") as config:
            config.write(linhas[0])
            config.write('\n')
            config.write(linhas[1])
            config.write('\n')
            config.write(linhas[2])
        savedirCR.configure(text=linhas[0])

    def fileselector_OR():
        linhas[1] = filedialog.askdirectory()
        with open("config.txt", "w") as config:
            config.write(linhas[0])
            config.write('\n')
            config.write(linhas[1])
            config.write('\n')
            config.write(linhas[2])
        savedirOR.configure(text=linhas[1])

    titleTB = customtkinter.CTkLabel(settings, text="Configurações", font=("Berlin Sans FB Demi", 24))
    titleTB.place(anchor="center", x=200, y=30)
    labelCR = customtkinter.CTkLabel(settings, text="Diretório de salvamento Tabela Bergerson", font=("Helvetica", 14))
    labelCR.place(anchor="center", x=200, y=70)
    savedirCR = customtkinter.CTkButton(settings, text="", command=fileselector_CR, width=325, height=30, fg_color="#242424", corner_radius=40, border_color="#485F72", border_width=1)
    savedirCR.place(anchor="center", x=200, y=100)
    labelOR = customtkinter.CTkLabel(settings, text="Diretório de salvamento Orçamentos", font=("Helvetica", 14))
    labelOR.place(anchor="center", x=200, y=150)
    savedirOR = customtkinter.CTkButton(settings, text="", command=fileselector_OR, width=325, height=30, fg_color="#242424", corner_radius=40, border_color="#485F72", border_width=1)
    savedirOR.place(anchor="center", x=200, y=180)
    destroysettings = customtkinter.CTkButton(settings, border_color="#485F72", border_width=1, text="< Voltar", command=destroy_settings, width=80, height=30, font=("Helvetica", 12, "italic"), corner_radius=40, fg_color="#242424")
    destroysettings.place(anchor="center", x=200, y=275)

    savedirCR.configure(text=linhas[0])
    savedirOR.configure(text=linhas[1])

def destroy_principal():
    principal.destroy()

principal = customtkinter.CTk()
principal.geometry("400x300+760+390")
principal.deiconify()
customtkinter.set_appearance_mode("dark")
principal.title("GScript for CIAF")
principal.resizable(False, False)

titleMain = customtkinter.CTkLabel(principal, text="GScript for CIAF", font=("Berlin Sans FB Demi", 28))
titleMain.place(anchor="center", x=200, y=40)
buttonCR = customtkinter.CTkButton(principal, text="Baixar Contas a Receber", command=CRWindow, width=325, height=35, font=("Helvetica", 14), corner_radius=40)
buttonCR.place(anchor="center", x=200, y=100)
buttonTAB = customtkinter.CTkButton(principal, text="Tabela Bergerson", command=TBWindow, width=325, height=35, font=("Helvetica", 14), corner_radius=40)
buttonTAB.place(anchor="center", x=200, y=145)
buttonORC = customtkinter.CTkButton(principal, text="Gerador de Orçamento", command=ORWindow, width=325, height=35, font=("Helvetica", 14), corner_radius=40)
buttonORC.place(anchor="center", x=200, y=190)
destroybtn = customtkinter.CTkButton(principal, border_color="#485F72", border_width=1, text="Encerrar", command=destroy_principal, width=80, height=30, font=("Helvetica", 12, "italic"), corner_radius=40, fg_color="#242424")
destroybtn.place(anchor="center", x=165, y=275)
settingsIMG = customtkinter.CTkImage(light_image=Image.open("img/settings.png"),
                                  dark_image=Image.open("img/settings.png"),
                                  size=(20, 20))
settingsbtn = customtkinter.CTkButton(principal, command=settings,border_color="#485F72", border_width=1, text="", width=20, height=30, image=settingsIMG, fg_color="#242424", corner_radius=40)
settingsbtn.place(anchor="center", x=240, y=275)

shutil.copy(r"C:/Ciaf/Dados/ORDEM-SERV.DBF", r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\ciaf-files")
shutil.copy(r"C:/Ciaf/Dados/ordem-serv.FPT", r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\ciaf-files")
clientes = []

principal.mainloop()