import tabula
import pandas
import pyautogui
import time
import tkinter
import customtkinter
import re

def ContasAPagar():
    tabela = tabula.read_pdf("t", pages="all")[0] #PDF Table read and transform to dataframe
    tabela.rename(columns=tabela.iloc[4], inplace = True)
    tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
    tabela = tabela.drop(0)
    tabela = tabela.drop(1)
    tabela = tabela.drop(2)
    tabela = tabela.drop(3)
    tabela = tabela.drop(4)

    pyautogui.click(984,1054) #Preparing CIAF to recieve the commands
    pyautogui.click(23,251)
    time.sleep(1)

    os = tabela["DOCTO:"].values #Transforming dataframe to array
    preco = tabela["R$ DEVIDO:"].values
    cont = 0

    for nos in os:
        pyautogui.click(1382,537) #Sequence of clicks
        time.sleep(0.5)
        pyautogui.typewrite(nos)
        time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click(1381,598)
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
        pyautogui.press('tab')
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


    titleTB = customtkinter.CTkLabel(windowTB, text="Gerar tabela Bergerson", font=("Berlin Sans FB Demi", 24))
    titleTB.pack(pady=10)
    labelCR = customtkinter.CTkLabel(windowTB, text="Insira as datas de intervalo da consulta:", font=("Helvetica", 14))
    labelCR.pack(pady=5)
    entryTB1 = customtkinter.CTkEntry(windowTB, placeholder_text="      data inicial", height=30, width=155, font=("Helvetica", 14,"italic"), corner_radius=40, text_color="white", state="normal")
    entryTB1.place(anchor="center", x=110, y=105)
    entryTB2 = customtkinter.CTkEntry(windowTB, placeholder_text="       data final", height=30, width=155, font=("Helvetica", 14, "italic"), corner_radius=40, text_color="white", state="normal")
    entryTB2.place(anchor="center", x=290, y=105)

    destroyCR = customtkinter.CTkButton(windowTB, command=destroy_tb, text="< Voltar", font=("Helvetica", 10, "italic"), width=80, height=30, fg_color="#242424", text_color="white", corner_radius=40)
    destroyCR.place(anchor="center", x=200, y=275)

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
destroybtn = customtkinter.CTkButton(principal, text="Encerrar", command=destroy_principal, width=80, height=30, font=("Helvetica", 12, "italic"), corner_radius=40, fg_color="#242424")
destroybtn.place(anchor="center", x=200, y=275)

principal.mainloop()