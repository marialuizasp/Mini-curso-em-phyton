import pandas as pd
import win32com.client as win32

#importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")


#vizualizar a base de dados
pd.set_option("display.max_columns", None)
print(tabela_vendas)


# faturamento por loja
faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento)


#quantidade de produtos vendidos por loja
produtos_vendidos = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print(produtos_vendidos)


#ticket médio por produto em cada loja
ticket_medio = (faturamento ["Valor Final"] / produtos_vendidos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)


#enviar um e-mail com relatorio
outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.to = "e-mail para quem você quer enviar"
mail.Subject = "Relatório de vendas por loja"
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={"Valor Final": "R${:,.2f}".format})}

<p>Quantidade vendida:</p>
{produtos_vendidos.to_html(formatters={"Quantidade": "R${:,.2f}".format})}

<p>Ticket medio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida estou a disposição</p>

<p>Att, </p>

<p>Seu Nome</p>
'''

mail.Send()
print("Email enviado") 