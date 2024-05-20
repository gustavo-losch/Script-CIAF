import tabula
import pyautogui
import time
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
    windowOR.geometry("1280x720")
    windowOR.resizable(False, False)

    def destroy_or():
        principal.deiconify()
        windowOR.destroy()

    top_frame = customtkinter.CTkFrame(master=windowOR, width=500, height=150, corner_radius=40)
    top_frame.grid(row=0,column=0, pady=(20,10), padx=(20,10), sticky="w")
    framelabel = customtkinter.CTkLabel(top_frame, text="Orçamento",fg_color="#2b2b2b", font=("Berlin Sans FB Demi", 20))
    framelabel.grid(row=0, column=0, padx=10,pady=10,sticky="w")
    n_orc = customtkinter.CTkSpinbo
    

    h_frame = customtkinter.CTkFrame(master=windowOR, width=280, height=150, corner_radius=40)
    h_frame.grid(row=1,column=0, pady=10, padx=20, sticky="w")

    m_frame = customtkinter.CTkFrame(master=windowOR, width=280, height=300, corner_radius=40)
    m_frame.grid(row=2,column=0, pady=10, padx=20, sticky="w")



    destroyOR = customtkinter.CTkButton(windowOR, command=destroy_or, text="< Voltar", font=("Helvetica", 10, "italic"), width=80, height=30, fg_color="#242424", text_color="white", corner_radius=40)
    destroyOR.grid(row=3,column=0, pady=5, padx=10)

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

principal.mainloop()