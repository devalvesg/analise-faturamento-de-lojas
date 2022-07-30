from tkinter import Pack
from turtle import right
from pyautogui import press, hotkey, click, write, position
from pyperclip import copy
from time import sleep
import mouse
import pandas as pd
import pyperclip

'''while True:
    print(mouse.get_position())
    sleep(2)'''

#entrando no arquivo e baixando planilha de dados excel

press('win')
write('google')
press('enter')
sleep(1)
copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
hotkey('ctrl', 'v')
press('enter')
sleep(3)
click(x=365, y=304, clicks=2)
sleep(2)
click(x=365, y=304, button='right')
sleep(1)
click(512, 652)

#computando dados planilha

tabela = pd.read_excel(r'C:\Users\gabri\Downloads/Vendas - Dez.xlsx')
quantidade = tabela['Quantidade'].sum()
faturamento = tabela['Valor Final'].sum()

#entrando no email


hotkey('ctrl','t')
sleep(1)
copy('https://outlook.live.com/mail/0/')
hotkey('ctrl', 'v')
press('enter')
sleep(5)
click(166, 171)
sleep(2)


#escrevendo destinatario e preparando assunto

write('gabrieldamasceno.bad@gmail.com')
press('enter')
press('tab')
sleep(2)
pyperclip.copy('RELATÓRIO DE VENDAS')
hotkey('ctrl', 'v')
sleep(1)
press('tab')

#escrevendo o email e 
texto = f'''Prezados, Bom dia
O faturamento foi de:{faturamento:,.2f}R$
A quantidade de vendas foi de:{quantidade:,} produtos'
Att.
Gabriel Alves'''
pyperclip.copy(texto)
hotkey('ctrl','v')
hotkey('ctrl','enter')