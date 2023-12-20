#  ---- biblioteca Pandas - instalar:
# pip install pandas

#  ---- o Pandas precisa de um pacote chamado openpyxl para ler arquivos excel: openpyxl
# pip install openpyxl --upgrade

#  ---- precisaremos uma biblioteca para integração do Python com o Windows
# pip install pywin32


import win32com.client as win32
import pandas as pd

# IMPORTAR A BASE DE DADOS
# VISULAZAR A BASE DE DADOS / instalar o openpyxl para leitura do excel pelo pandas
# o pandas vai ler o arquivo em excel, e armazenar esse arquivo excel dentro da nossa tabela_vendas
tabela_vendas = pd.read_excel('Vendas.xlsx')

# ajustar a vizualização da nossa base de dados
pd.set_option('display.max_columns', None)  # mostre TODOS os dados da tabela (colunas), sem limites de colunas

# sempre que for pra filtar uma coluna: tabela_vendas [['...', '...']] ou tabela_vendas.groupby('...').sum() \
# uma lista tem q ficar dentro de cochetes
# ... sum() vai somar os valores da coluna em questão
# ... grupoby irá fazer o agrupamento de cada quesito
# ... [[]] filtar várias colunas e [] filtra uma coluna

#                                     --------FATURAMENTO POR LOJA---------
# filtrar somente as colunas de loja e faturamento
# groupby() para agrupar todas lojas (cada loja aparecendo uma única vez),  .sum() para o somatório da coluna faturamento.
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('_' * 50)

#                         ---------QUANTIDADE DE PRODUTOS DE PRODUTOS VENDIDOS POR LOJA-------
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('_' * 50)

#                               ---------TICKET MÉDIO POR PRODUTO EM  CADA LOJA--------
# SEMPRE QUE FAÇO UMA OPERAÇAO ENTRE COLUNAS (terei como resultado uma série de dados, ao invés de uma ''tabela'') Dessa forma, para TRANSFORMAR esse resultado EM TABELA, e que fique bonitinha, EU COLOCO um .to_frame NO FINAL.
# como zero é um número, n precisamos utilizar aspas, pois o Python entende os números
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um e-mail com o relatório pelo outlook
# instalar pywin32 para integrar python com e-mail outlook:  pip install pywin32 NO TERMINAL
# código abaixo se encontra disponível na net
# import win32com.client as win32
# import win32com.client as win32
outlook = win32.Dispatch('outlook.application') # variável do outlook (se conecta ao outlook do computador)
mail = outlook.CreateItem(0)  # uma variável item do outlook; cria um e-mail
mail.To = 'emailtest@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja' # assunto do e-mail

# mail.HTMLBody: é opcional. O corpo do nosso e-mail é um html, dessa forma, ele  pode ser todo personalizado. Com isso, temos q trnsformar essa tabela, numa tabela html (tabela bonitinha personalizada), e para isso utilizo o to_html()
# o f na frente de um texto, significa que posso ter chaves, onde dentro dessas, posso passar variáveis do meu código (formatar essas chaves com variáveis)
# # formatters={'Valor Final': 'R${:,.2f}'.format  o formatters vai formatar nosso número={quem queremos modificar: o que será modificado}
# toda formatação começa com 2 pontos: a vírgula é o separador de milhar; o ponto, separador decimal; e 2f, quantas casas decimais (2 float), ou seja, duas casas decimais.
mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por Loja:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})} 

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:<p/>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>

<p>Tércio Matos</p>
'''
mail.Send()
