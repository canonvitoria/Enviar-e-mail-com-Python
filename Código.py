import win32com.client as win32

 #Criar a integração com o outlook/ter ele instalado e configurado

outlook = win32.Dispatch('outlook.aplication')

#Criar email
email = outlook.CreatItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento/qtde_produtos

#configurar as informações do seu e-mail
email.To = "EMAIL@gmail.com"
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p>Olá, aqui é o código Python</p>

<p>O faturamento da loja foi de R${faturamento}<\p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Média foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
"""

#Anexo = "ENDEREÇO DO LOCAR DO ARQUIVO"
#Por exemplo ->> C://Users/joaop/Downloads/arquivo.xlsx
#email.Attachments.Add(ANEXO)

email.Send()
print("Email Enviado")