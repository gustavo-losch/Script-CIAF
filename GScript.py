import tabula
import pyautogui
import time
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
from PIL import Image
import shutil
import csv

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
            time.sleep(5)

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
            time.sleep(1)

            os = tabela["DOCTO:"].values #Transforming dataframe to array
            preco = tabela["R$ DEVIDO:"].values
            cont = 0

            for nos in os:
                pyautogui.click(1382,537) #Localizar
                time.sleep(0.5)
                pyautogui.typewrite(nos) #Digitar OS
                time.sleep(0.5)
                pyautogui.press('enter') #Tabular
                pyautogui.press('enter') #Confirmar Localizar
                time.sleep(0.5)
                pyautogui.press('tab') #Tabular para não baixar
                time.sleep(0.5)
                pyautogui.press('enter') #Confirmar não biaxar
                time.sleep(0.5)
                pyautogui.click(1381,598) #Clicar em Baixar
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('tab') #data
                time.sleep(0.5)
                pyautogui.typewrite(dataCR)
                time.sleep(0.5)
                ##pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('d')
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.typewrite(tabela.iloc[cont,2])
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(0.5)
                pyautogui.click(953,590)
                time.sleep(1)

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
    principal.iconify()
    windowOR.title("Orçamento")
    windowOR.geometry("1600x720")
    windowOR.resizable(True,True)

    global clientes
    global nomes_clientes  
    global telefone_clientes  
    global cpf_clientes
    global orcamentos
    orcamentos = []

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

    def novo_orcamento():
        global n_orc
        with open("config.txt", "r") as config:
            linhas = config.read().splitlines()

        n_orc = int(linhas[2])
        n_orc += 1
        linhas[2] = str(n_orc)

        with open("config.txt", "w") as config:
            for linha in linhas:
                config.write(linha)
                config.write('\n')
        
        #resetar campos

    def exportar_orcamento():
        def dados_horas():
            global n_orc
            def adicionar_servico(df, df_horas, n_orc, servico):
                if df.loc[n_orc, servico] != 0:
                    df_horas.loc[len(df_horas),"Serviço"] = servico.capitalize()
                    if df.loc[n_orc,"time_format"] == "Minutos":
                        df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"] = format((df.loc[n_orc,servico])/60, ".2f")
                        df_horas.loc[len(df_horas)-1,"Valor"] = format(float((df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"])*df.loc[n_orc,"preco_hora"]), ".2f")
                    else:
                        df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"] = (df.loc[n_orc,servico])
                        df_horas.loc[len(df_horas)-1,"Valor"] = format((df_horas.loc[len(df_horas)-1,"Horas Trabalhadas"])*df.loc[n_orc,"preco_hora"], ".2f")

            with open(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\config.txt", "r") as config:
                linhas = config.read().splitlines()
                n_orc = int(linhas[2])

            df = pd.read_csv(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\orcamentos.csv", encoding="ISO-8859-1")
            df = df.fillna(0)

            servicos = ["prototipagem", "desenho", "molde", "fundicao", "montagem", "acabamentos", "polimento", "limpeza", "cravacao"]

            df_horas = pd.DataFrame(columns=["Serviço", "Horas Trabalhadas", "Valor"])

            for servico in servicos:
                adicionar_servico(df, df_horas, n_orc, servico)
            
            

            table_data = [df_horas.columns.to_list()] + df_horas.values.tolist()
            return table_data
        

        tabela_hora = dados_horas



        
    def salvar_orcamento():

        global n_orc
        with open("config.txt", "r") as config:
            linhas = config.read().splitlines()

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
                    "preco_hora":preco_hora}
        orcamentos.append(orcamento)

        with open("orcamentos.csv", mode='w', newline='') as orc:
            writer = csv.DictWriter(orc, fieldnames=["n_orc","data_emissao","data_validade","nome_cli","descricao","time_format","prototipagem","desenho","molde","fundicao","montagem","acabamentos","polimento","limpeza","cravacao","ouro1k","ouro750","ouro_branco","pedras","prata","rodio","servicos_terceiros","cotacao","preco_hora"])
            writer.writeheader()
            for orcamento in orcamentos:
                writer.writerow(orcamento)
        
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

    def destroy_or():
        principal.deiconify()
        windowOR.destroy()

    top_frame = customtkinter.CTkFrame(master=windowOR, width=375, height=150, corner_radius=40)
    top_frame.place(y=30, x=30, anchor="nw")
    title_label = customtkinter.CTkButton(top_frame, text="Orçamento",width=200, height=50, fg_color="#242424", border_color="#1f6aa5", border_width=1, corner_radius=40, font=("Berlin Sans FB Demi", 22), hover=False).place(y=40, x=120, anchor="center")
    n_or = customtkinter.CTkButton(top_frame, text="0001",width=80, height=50, corner_radius=40, fg_color="#1f6aa5",font=("Arial", 20, "bold"), hover=False).place(y=40, x=306, anchor="center")
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

    save_btn = customtkinter.CTkButton(windowOR, text="Salvar",width=250, height=40, command=salvar_orcamento, font=("Berlin Sans FB Demi", 22), corner_radius=40)
    save_btn.place(anchor="center", x=1100, y=680)

    destroyOR = customtkinter.CTkButton(windowOR, command=destroy_or, text="< Voltar", font=("Helvetica", 10, "italic"), width=80, height=30, fg_color="#242424", text_color="white", corner_radius=40)
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
principal.deiconify()
customtkinter.set_appearance_mode("dark")
principal.title("GScript for CIAF")
principal.geometry("400x300")
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
settingsIMG = customtkinter.CTkImage(light_image=Image.open(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\settings.png"),
                                  dark_image=Image.open(r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\settings.png"),
                                  size=(20, 20))
settingsbtn = customtkinter.CTkButton(principal, command=settings,border_color="#485F72", border_width=1, text="", width=20, height=30, image=settingsIMG, fg_color="#242424", corner_radius=40)
settingsbtn.place(anchor="center", x=240, y=275)

shutil.copy(r"C:/Ciaf/Dados/ORDEM-SERV.DBF", r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\ciaf-files")
shutil.copy(r"C:/Ciaf/Dados/ordem-serv.FPT", r"C:\Users\Gustavo Losch\Documents\Repositórios\Script-CIAF\ciaf-files")
clientes = []

principal.mainloop()