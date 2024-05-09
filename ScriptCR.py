import tabula
import pandas
import pyautogui
import time

tabela = tabula.read_pdf("Path.pdf", pages="all")[0] #PDF Table read and transform to dataframe
tabela.rename(columns=tabela.iloc[4], inplace = True)
tabela = tabela[['DOCTO:', 'VENCIMENTO:', 'R$ DEVIDO:']]
tabela = tabela.drop(0)
tabela = tabela.drop(1)
tabela = tabela.drop(2)
tabela = tabela.drop(3)
tabela = tabela.drop(4)


pyautogui.click(1239,1051) #Preparing CIAF to recieve the commands
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

    print(nos)
    print(tabela.iloc[cont,2])

    cont = cont + 1 #Line skip

    

