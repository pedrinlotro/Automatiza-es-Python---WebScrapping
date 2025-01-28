import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import pandas as pd

def create_weekly_excel(weekly_current, weekly_previous, weekly_variation):
    """Cria um arquivo Excel com os dados da variação semanal do alumínio."""
    data = {
        "Data": [datetime.now().strftime("%d-%m-%Y")],
        "Preço Atual (Semanal)": [weekly_current],
        "Preço Anterior (Semanal)": [weekly_previous],
        "Variação Semanal (%)": [weekly_variation]
    }
    df = pd.DataFrame(data)
    file_path = "aluminum_weekly_price.xlsx"
    df.to_excel(file_path, index=False)
    print(f"Arquivo Excel criado: {file_path}")
    return file_path

def send_email_with_sendgrid(receiver_emails, file_path):
    """Envia um e-mail com o arquivo Excel anexado usando SendGrid."""
    smtp_server = "smtp.sendgrid.net"
    port = 587
    sender_email = "xxx"
    sendgrid_api_key = "xxx"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receiver_emails)
    msg['Subject'] = "Variação Semanal do Alumínio"

    body = "Segue em anexo o arquivo com a variação semanal do alumínio."
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={file_path}")
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login("apikey", sendgrid_api_key)
        server.send_message(msg)
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")
    finally:
        server.quit()

# Simula dados de entrada da variação semanal
weekly_current = 2631.75
weekly_previous = 2595.20
weekly_variation = 1.41  # Em %

# Criar o arquivo Excel
file_path = create_weekly_excel(weekly_current, weekly_previous, weekly_variation)

# Lista de destinatários
receiver_emails = [
    "xxx"
]

# Enviar o e-mail
send_email_with_sendgrid(receiver_emails, file_path)
