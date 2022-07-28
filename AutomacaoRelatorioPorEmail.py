import pyautogui
import pyperclip
import time
import pandas as pd
import openpyxl
import numpy

#pyautogui.click -> clicar com o mouse
#pyautogui.write -> escrever um texto
#pyautogui.press -> pressionar um botão
#pyautogui.hotkey() -> "ctrl", "v"

pyautogui.PAUSE = 0.7

# Passo 1: Entrar no sistema (link)
#pyautogui.hotkey("ctrl", "t")
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# site ta carregando
time.sleep(4)

# Passo 2: Navegar no sistema e encontrar a base de dados (entrar na pasta Exportar)
pyautogui.click(x=450, y=318, clicks=2)
time.sleep(1)

# Passo 3: Download da base de dados
pyautogui.click(x=450, y=318, clicks=1) # Clicou no arquivo
pyautogui.click(x=1680, y=197) # Clicou nos 3 pontinhos
pyautogui.click(x=1398, y=703) # Fazer download

#ver a posicao dos frames x e y
#time.sleep(5)
#print(pyautogui.position())

# Passo 4: Calcular os indicadores (faturamento, quantidade de produtos)
tabela = pd.read_excel(r"C:\Users\leoro\Downloads\Vendas - Dez.xlsx")
print(tabela)

quantidade = tabela["Quantidade"].sum()  # Soma da coluna quantidade da tabela

faturamento = tabela["Valor Final"].sum()  # Soma da coluna valor final da tabela

print(quantidade)
print(faturamento)

# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")  #Nova aba
pyperclip.copy("mail.google.com/mail/u/0/#inbox")  #Copiando Link
pyautogui.hotkey("ctrl", "v") #COlando o link
pyautogui.press("enter") #Entrando no site
time.sleep(7)

# Passo 6: Mandar um email para a diretoria com os indicadores
#clicar no botao +
pyautogui.click(x=81, y=213)
time.sleep(2)

#Escrever o destinatario
pyautogui.write("leonardo@rodaelli.com.br")
pyautogui.press("tab")  #selecionar email
pyautogui.press("tab")  #passar pro campo do assunto

# escrever o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # Passar para o corpo do email

# escrever o corpo do email
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A Quantidade de produtos foi de: {quantidade:,}

Abs
Leo Rodaelli
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# Enviar o email
pyautogui.hotkey("ctrl", "enter")

