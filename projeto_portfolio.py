import pandas as pd
import win32com.client as win32

# IMPORTAR A BASE DE DADOS
tabela_vendas = pd.read_excel('Vendas.xlsx')


# VISUALIZAR A BASE DE DADOS
pd.set_option('display.max_columns', None)


# FATURAMENTO POR LOJA
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)


# QUANTIDADES DE PRODUTO VENDIDO POR LOJA
quantidade_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_produtos)


# TICKET MÉDIO POR LOJA
ticket = (faturamento['Valor Final']/quantidade_produtos['Quantidade']).to_frame()
print(ticket)

# ENVIAR EMAIL COM RELATÓRIO
outlook = win32.Dispatch('outlook.application')
mail = outlook.Createitem(0)
mail.To = 'gabrieldamasceno.bad@gmail.com'
mail.Subject = 'Relatório de vendas'
mail.HTMLBody = f'''
<p>Prezados,
<p>Segue o relatório de vendas por cada loja desse mês</p>

<p>Faturamento:</p>
{faturamento.to_html()}

<p>Quantidade:</p>
{quantidade_produtos.to_html()}

<p>Ticket médio dos produtos por cada loja:</p>
{ticket.to_html()}

<p>Qualquer dúvida estou a disposição</p>

<p>Att.</p>
<p>Gabriel</p>
'''
mail.Send()
print('Hellow World')