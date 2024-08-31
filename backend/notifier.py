import os
from dotenv import load_dotenv
from twilio.rest import Client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename


# Carregue as variáveis do arquivo .env
load_dotenv()

# Obtenha as variáveis de ambiente
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')
TWILIO_ACCOUNT_SID = os.getenv('TWILIO_ACCOUNT_SID')
TWILIO_AUTH_TOKEN = os.getenv('TWILIO_AUTH_TOKEN')
TWILIO_PHONE_NUMBER = os.getenv('TWILIO_PHONE_NUMBER')
SMS_RECEIVER_NUMBER = os.getenv('SMS_RECEIVER_NUMBER')

def send_email(subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))
    
    if attachment_path:
        try:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEApplication(attachment.read(), Name=basename(attachment_path))
                part['Content-Disposition'] = f'attachment; filename="{basename(attachment_path)}"'
                msg.attach(part)
        except Exception as e:
            print(f"Erro ao anexar o arquivo: {e}")

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, EMAIL_RECEIVER, msg.as_string())
        print("Email enviado com sucesso!")
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar e-mail: {e}")
def send_sms(to, body):
    client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    try:
        message = client.messages.create(
            from_=TWILIO_PHONE_NUMBER,
            body=body,
            to=to
        )
        print("SMS enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar SMS: {e}")

if __name__ == "__main__":
    send_email('Assunto do E-mail', 'Corpo do e-mail')
    send_sms(TWILIO_PHONE_NUMBER, 'Mensagem de teste')