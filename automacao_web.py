from dataclasses import replace
from msilib.schema import tables
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import pandas as pd

#PEGANDO COTAÇÕES
nav = webdriver.Chrome()
nav.get('https://www.google.com.br/')


#Usando o xpath e pegando cotação dolar

nav.find_element('xpath',
 '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dolar')  
nav.find_element('xpath',
 '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER) 

dolar = nav.find_element('xpath', 
'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')


#pegando cotação euro

nav.find_element('xpath', '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input').send_keys(Keys.BACKSPACE*20)
nav.find_element('xpath', '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input').send_keys('cotação do euro')
nav.find_element('xpath', '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input').send_keys(Keys.ENTER)

euro = nav.find_element('xpath', 
'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')


#cotação ouro

nav.get('https://www.melhorcambio.com/ouro-hoje')

ouro = nav.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')
ouro = ouro.replace(',', '.')

nav.quit()

#base de dados
tabela = pd.read_excel('Produtos.xlsx')

#recalcular preços

tabela.loc[tabela['Moeda']=='Dólar','Cotação'] = float(dolar)
tabela.loc[tabela['Moeda']=='Euro','Cotação'] = float(euro)
tabela.loc[tabela['Moeda']=='Ouro','Cotação'] = float(ouro)

#atualizando preços

tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

#exportando para excel

tabela.to_excel('Produtos Novo.xlsx', index=False)