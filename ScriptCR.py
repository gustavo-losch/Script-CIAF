import tabula
import pandas
import pyautogui
import time
import tkinter
import customtkinter

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
        if len(dataCR) == 10:
            print("Rodar")
            entryCR.configure(state="disabled")
            erroCR.configure(text="")
            runCR.configure(text="O código ira rodar em 5 segundos.")
            time.sleep(5)

            tabela = tabula.read_pdf("Loveggy.pdf", pages="all")[0] #PDF Table read and transform to dataframe (insert PDF path)
            tabela.rename(columns=tabela.iloc[4], inplace = True)
            tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
            tabela = tabela.drop(0)
            tabela = tabela.drop(1)
            tabela = tabela.drop(2)
            tabela = tabela.drop(3)
            tabela = tabela.drop(4)

            pyautogui.click(984,1055) #Preparing CIAF to recieve the commands
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

        else:
            print("errado")
            runCR.configure(text="")
            erroCR.configure(text="Insira a data novamente.")

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
principal.grid_columnconfigure((0), weight=1)

buttonCR = customtkinter.CTkButton(principal, text="Baixar Contas a Receber", command=CRWindow)
buttonCR.grid(row=0, column=0, padx=20, pady=20, sticky="ew", columnspan=2)

principal.mainloop()