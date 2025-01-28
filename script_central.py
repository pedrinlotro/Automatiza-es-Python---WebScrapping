import schedule
import time
import os
from datetime import datetime
import pytz

# Define o fuso horário de Brasília
tz_brasilia = pytz.timezone("America/Sao_Paulo")

def run_daily_script():
    """Função para rodar o script diário."""
    os.system(r"C:\Users\pedro.ramos\AppData\Local\Programs\Python\Python313\python.exe C:\Users\pedro.ramos\Desktop\auto_script\script_alum.py")
    now = datetime.now(tz_brasilia).strftime('%d/%m/%Y %H:%M:%S')
    print(f"Script diário executado com sucesso em {now} horário de Brasília!")

def run_weekly_script():
    """Função para rodar o script semanal."""
    os.system(r"C:\Users\pedro.ramos\AppData\Local\Programs\Python\Python313\python.exe C:\Users\pedro.ramos\Desktop\auto_script\script_semanal.py")
    now = datetime.now(tz_brasilia).strftime('%d/%m/%Y %H:%M:%S')
    print(f"Script semanal executado com sucesso em {now} horário de Brasília!")

# Agendar script diário para rodar todos os dias às 7:00 no horário de Brasília
schedule.every().day.at("11:00").do(run_daily_script)

# Agendar script semanal para rodar todas as sextas-feiras às 07:01 no horário de Brasília
schedule.every().friday.at("11:01").do(run_weekly_script)

# Mensagem informando que o agendador está ativo
print("Agendador iniciado. Aguardando execuções diárias e semanais...")

# Loop para manter o agendador ativo
while True:
    schedule.run_pending()
    time.sleep(60)  # Verifica a cada 60 segundos para reduzir o uso de CPU
