import tabula
import pandas
import pyautogui
import time

tabela = tabula.read_pdf("Extrato Cecilia.pdf", pages="all")[0]
tabela.rename(columns=tabela.iloc[4], inplace = True)
tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
tabela = tabela.drop(0)
tabela = tabela.drop(1)
tabela = tabela.drop(2)
tabela = tabela.drop(3)
tabela = tabela.drop(4)


pyautogui.click(1239,1051)
pyautogui.click(23,251)
time.sleep(1)


os = tabela["DOCTO:"].values
preco = tabela["R$ DEVIDO:"].values
cont = 0

for nos in os:
    pyautogui.click(1382,537)
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
    pyautogui.click(1380,598)
    time.sleep(0.5)
    pyautogui.click(991,592)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('tab')
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
    time.sleep(1)

    print(nos)
    print(tabela.iloc[cont,2])
    cont = cont + 1

    

