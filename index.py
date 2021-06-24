import pyautogui
import time
import pyperclip


pyautogui.PAUSE = 1
pyautogui.alert("Vai começar, aperte OK e não mexa em nada")

pyautogui.press("winleft")
pyautogui.write("Navegador Opera GX")
pyautogui.press("enter")

#Chegar até a apostila
pyautogui.hotkey("ctrl", "t")
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#Chegar à tabela
time.sleep(3)
pyautogui.click(x=452, y=264, clicks=2)

#Download tabela
pyautogui.PAUSE = 1
pyautogui.click(438, 385)
pyautogui.click(1709, 159)
pyautogui.click(1545, 532)
time.sleep(3)

#tabela soma e etc
import pandas as pd
tabela = pd.read_excel("D:\Meus Documentos\Downloads/Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
import pyautogui
import time
import pyperclip
pyautogui.PAUSE = 1

#enviar e-mail
pyautogui.hotkey("ctrl", "t")


pyautogui.write("https://mail.google.com")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(143, 176)
pyautogui.write("jpbcs2005@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", 'v')
pyautogui.press("tab")
nome = ("João Pedro")
texto_email = f"""
Prezados, bom dia

O faturamento foi de : R${faturamento:,.2f}
A quantidade de produtos foi : {quantidade}

Abs
JP
"""
pyautogui.write(texto_email)

pyautogui.hotkey("ctrl", "enter")
pyautogui.PAUSE = 0.4
pyautogui.hotkey("ctrl", "w")
pyautogui.hotkey("ctrl", "w")
pyautogui.hotkey("ctrl", "w")
pyautogui.hotkey("alt", "f4")
pyautogui.alert("Fim da Automação. Seu computador já voltou a ser seu")
