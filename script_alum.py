import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Verifica se hoje é final de semana (sábado = 5, domingo = 6)
today = datetime.now().weekday()
if today in [5, 6]:
    print("Hoje é final de semana. O script não será executado.")
    exit()

def fetch_prices():
    """Extrai os preços do alumínio e dólar do site Shock Metais exatamente como aparecem."""
    url = "https://shockmetais.com.br/lme"

    # Configuração para modo headless (sem interface gráfica)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        driver.implicitly_wait(10)

        # Formata a data de hoje no formato correto da tabela (ex: "27/Jan")
        today_str = datetime.now().strftime("%d/%b").lstrip("0")

        # Localiza as linhas da tabela
        rows = driver.find_elements(By.XPATH, '//*[@id="tabelalme"]/div/div[3]/div/div/table/tbody/tr')

        for row in rows:
            columns = row.find_elements(By.TAG_NAME, 'td')
            if columns[0].text.strip() == today_str:
                aluminum_price_usd = columns[3].text.strip()  # Captura exatamente como está no site
                dollar_price = columns[7].text.strip()  # Captura exatamente como está no site

                print(f"Data: {today_str}, Alumínio: {aluminum_price_usd} USD/t, Dólar: {dollar_price} BRL")
                return aluminum_price_usd, dollar_price

        print("Data de hoje não encontrada na tabela.")
        return None, None

    except Exception as e:
        print(f"Erro ao extrair dados: {e}")
        return None, None
    finally:
        driver.quit()

def create_or_update_excel(aluminum_price_usd, dollar_price):
    """Atualiza ou cria um arquivo Excel com os dados extraídos, preservando o formato correto."""

    # Calcula o preço do alumínio em BRL/t mantendo a formatação correta
    aluminum_price_brl = f"{float(aluminum_price_usd.replace('.', '').replace(',', '.')) * float(dollar_price.replace(',', '.')):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    today_str = datetime.now().strftime("%d-%m-%Y")

    new_data = {
        "Data": today_str,
        "Preço Alumínio (USD/t)": aluminum_price_usd,  # Mantém o formato original
        "Preço do Dólar (BRL)": dollar_price,  # Mantém o formato original
        "Preço Alumínio (BRL/t)": aluminum_price_brl,
    }

    file_path = "aluminum_price.xlsx"
    
    try:
        existing_df = pd.read_excel(file_path, dtype={"Data": str})
        if today_str in existing_df["Data"].values:
            existing_df.loc[existing_df["Data"] == today_str, ["Preço Alumínio (USD/t)", "Preço do Dólar (BRL)", "Preço Alumínio (BRL/t)"]] = [
                new_data["Preço Alumínio (USD/t)"], new_data["Preço do Dólar (BRL)"], new_data["Preço Alumínio (BRL/t)"]
            ]
            print("Dados do dia atual atualizados.")
        else:
            new_row = pd.DataFrame([new_data])
            existing_df = pd.concat([existing_df, new_row], ignore_index=True)
            print("Nova entrada adicionada.")
    except FileNotFoundError:
        existing_df = pd.DataFrame([new_data])
        print("Arquivo criado com nova entrada.")

    existing_df.to_excel(file_path, index=False)
    print(f"Arquivo Excel atualizado: {file_path}")
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
    msg['Subject'] = "Preço do Alumínio Convertido"

    body = "Segue em anexo o arquivo atualizado com o preço do alumínio convertido para reais."
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

# Obter os preços do site
aluminum_price_usd, dollar_price = fetch_prices()

if aluminum_price_usd and dollar_price:
    # Atualizar ou criar o arquivo Excel
    file_path = create_or_update_excel(aluminum_price_usd, dollar_price)

    # Lista de destinatários
    receiver_emails = ["xxx"]

    # Enviar o e-mail
    send_email_with_sendgrid(receiver_emails, file_path)
else:
    print("Erro ao obter os preços. O script não será executado.")
