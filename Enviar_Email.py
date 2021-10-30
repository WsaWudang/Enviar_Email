import win32com.client as win32

#integração com outlook
outlook = win32.Dispatch('outlook.application')

#Criar email
email = outlook.CreateItem(0)

#caso tenha anexo
anexo = "local do arquivo que deseja enviar"
email.Attachments.Add(anexo)

#Para quem deseja enviar
email.to = "e-mails para quem deseja enviar;" 

#Titulo do email
email.Subject = "Teste Python"

#corpo do e-mail
email.HTMLBody = """
<p>Olá, estou enviando esse e-mail para teste de automação. </p>


<p>atenciosamente,</p>

<p>Wesley Almeida</p>
"""

#enviar email
email.Send()
print("E-mail enviado!")