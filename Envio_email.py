#Importar as Bibliotecas

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Importar base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabela_vendas)

#Faturamento por loja

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#Quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

#Ticket médio por produto em cada loja

ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#Enviar um relatorio via email

smtp_server = 'smtp.gmail.com'
smtp_port = 587
email = 'christopherkevin78@gmail.com'
senha = 'rnga abtr bdjx naxh'
Destinatario = 'christopherkevin78@gmail.com'
Assunto = 'Relatório de vendas por loja'
email_corpo = f'''
<p>Prezados,</p>

<p>Segue relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade:<p>
{quantidade.to_html()}

<p>Tickets médios dos produtos em cada loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Kevin</p>
'''
mensagem = MIMEMultipart()
mensagem['From'] = email
mensagem['To'] = Destinatario
mensagem['Subject'] = Assunto
mensagem.attach(MIMEText(email_corpo, 'html'))

try:
    servidor = smtplib.SMTP(smtp_server, smtp_port)
    servidor.starttls()
    servidor.login(email, senha)
    servidor.sendmail(email, Destinatario, mensagem.as_string())
    
    print('Email Enviado')

except Exception as e:
    print('Erro ao enviar o email: ', str(e))
    
finally:
    servidor.quit()

