import smtplib
from email.message import EmailMessage

def enviar_email(destinatario, assunto, corpo, anexo, remetente, senha):
    try:
        msg = EmailMessage()
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = destinatario
        msg.set_content(corpo)

        with open(anexo, 'rb') as f:
            file_data = f.read()
            file_name = anexo
        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(remetente, senha)
            smtp.send_message(msg)

        return True, "📨 E-mail enviado com sucesso!"
    except Exception as e:
        return False, f"❌ Erro ao enviar e-mail: {e}"
