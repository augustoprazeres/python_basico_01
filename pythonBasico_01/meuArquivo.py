import pandas as pd
import win32com.client as win32

#importar base de dados
tabela_vendas = pd.read_excel('./BaseDados/Vendas.xlsx')


#visualizar a base de dados
pd.set_option('display.max_columns', None) #comando para mostrar todas as colunas

print(tabela_vendas)
#print(tabela_vendas[['ID Loja', 'Valor Final']])
#print (tabela_vendas.groupby('ID Loja').sum())
print('--'*50)

#faturmaneot por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() #agrupa por loja, e soma as demais
#mostrar_tabela_agrupada_media = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').mean()
print('Faturamento total por loja: ', faturamento)
#print('Média total por loja: ',mostrar_tabela_agrupada_media)

print('--'*50)

#qtde de produtos vendidos por loja

quantidade_produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print('A quantidade de produtos vendidos por loja foi: ', quantidade_produtos_vendidos)
print('--'*50)
#ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade_produtos_vendidos['Quantidade']).to_frame()
print('O Ticket médio é:', ticket_medio)

ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#enviar um e-mail com relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pythonaugusto@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados</p>,

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format })}


<p>Quantidade vendida:</p>
{quantidade_produtos_vendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou a disposição.</p>
<p>Att.</p>
<p>Augusto Prazeres</p>
'''
mail.Send()

print('E-mail enviado')





