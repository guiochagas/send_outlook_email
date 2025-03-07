import win32com.client
from datetime import date, timedelta

hoje = date.today().strftime("%d/%m/%Y")
ontem = (date.today() - timedelta(days=1)).strftime("%d/%m/%Y")

def criar_email():
    try:
        # Criar objeto do Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)  # Criar novo e-mail

        # Escreva o e-mail em formato HTML.
        corpo_email = f"""
        <html>
        <body>
            <p>Bom dia,</p>
            <p>Este é o corpo do meu e-mail sendo escrito em formato HTML e é 100% personalizável dentro das possibilidades do HTML.</p>
            <br>
        </body>
        </html>
        """

        # Configurar destinatários e conteúdo
        email.To = "guilhermecchagas1999@gmail.com"  # E-mails de destino
        email.CC = "guilhermecchagas1999@gmail.com"  # E-mail em cópia
        email.Subject = f"E-MAIL AUTOMATIZADO FEITO DIA {hoje}"
        email.HTMLBody = corpo_email

        # Para adicionar um anexo, descomente a linha abaixo.
        # email.Attachments.Add(f"C:/CAMINHO/DO/ARQUIVO")

        # Exibir o e-mail antes de enviar (para edição manual)
        email.Display()  

        # Para enviar diretamente, descomente a linha abaixo (CUIDADO: Isso envia sem abrir!)
        # email.Send()

        print("📨 E-mail gerado com sucesso!")

    except Exception as e:
        print(f"⚠️ Erro ao criar o e-mail: {e}")

# Executar a função
criar_email()
