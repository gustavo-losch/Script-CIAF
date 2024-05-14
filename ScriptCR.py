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
    principal.iconify()
    windowCR = customtkinter.CTkToplevel(principal)
    windowCR.title("Contas a Receber")
    windowCR.geometry("400x250")
    windowCR.resizable(False, False)

    def iniciar_processo_cr():
        dataCR = entryCR.get()
        dataCR = re.sub('[/]', '', dataCR)
        if len(dataCR) == 8:
            print("Rodar")
            erroCR.configure(text="")
            runCR.configure(text="O código ira rodar em 5 segundos.")
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
            runCR.configure(text="")
            erroCR.configure(text="Insira a data novamente. (Data Inválida)")

    titleCR = customtkinter.CTkLabel(windowCR, text="Baixar Contas a Receber", font=("Berlin Sans FB Demi", 24))
    titleCR.pack(pady=10)
    labelCR = customtkinter.CTkLabel(windowCR, text="Insira a data de pagamento das contas:", font=("Helvetica", 14))
    labelCR.pack(pady=5)
    entryCR = customtkinter.CTkEntry(windowCR, placeholder_text="                         dia/mes/ano", height=30, width=300, font=("Helvetica", 14), corner_radius=40, text_color="white", state="normal")
    entryCR.pack(pady=0)
    erroCR = customtkinter.CTkLabel(windowCR, text="", font=("Helvetica", 10), text_color="red")
    erroCR.pack(pady=1)
    runCR = customtkinter.CTkLabel(windowCR, text="", font=("Helvetica", 10), text_color="green")
    runCR.pack(pady=1)

    startCR = customtkinter.CTkButton(windowCR, command=iniciar_processo_cr, text="Iniciar Processo", width=300, height=40, font=("Berlin Sans FB Demi", 18), corner_radius=40)
    startCR.pack(pady=10)

principal = customtkinter.CTk()
principal.deiconify()
customtkinter.set_appearance_mode("dark")
principal.title("GScript for CIAF")
principal.geometry("400x150")

titleMain = customtkinter.CTkLabel(principal, text="GScript for CIAF", font=("Berlin Sans FB Demi", 28))
titleMain.pack(pady=10)
buttonCR = customtkinter.CTkButton(principal, text="Baixar Contas a Receber", command=CRWindow, width=300)
buttonCR.pack(pady=20)

principal.mainloop()