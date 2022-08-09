import pandas as pd
import win32com.client as win32

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)
# quantidade de produtos vendidos por loja
print('-' * 50)
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
# ticket médio por produto em cada loja
print('-' * 50)
ticket = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})
print(ticket)
# enviar um email com o relatório


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'viniciusbenedito004@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>


<p>Faturamento:</p> 
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format })}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos Produtos em cada Loja:</p>
{ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format })}

<p>Qualquer dúvida estou à disposição.</p>

 <p>Att</p>
 <p>Vinicius Benedito</p>
'''

mail.Send()

print('Email Enviado')
