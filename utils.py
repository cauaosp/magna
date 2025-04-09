import smtplib
import subprocess
import openpyxl
import json
import schedule
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from scapy.all import IP, ICMP, sr1, conf, get_if_addr
import requests
from types import SimpleNamespace
import json

def load_json_as_object():
    with open("auth/auth.json", "r") as file:
        return json.load(file, object_hook=lambda d: SimpleNamespace(**d))

auth = load_json_as_object()

def read_sheet(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    header = [cell.value for cell in ws[1]]

    excel_list = []

    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        linha_dict = {}
        for i, valor in enumerate(row):
            if valor is not None:
                linha_dict[header[i]] = valor
        if linha_dict.get("IP") is not None:
            excel_list.append(linha_dict)
    return excel_list

def load_email_config():
    with open('auth/auth.json', 'r') as f:
        data = json.load(f)
    
    email_config = data["email_smtp"]
    return email_config

def html_content(broke_printers):
    table = """
    <html>
        <head>
            <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
        </head>
        <body>
            <p><strong>Problemas com as seguintes impressoras:</strong></p>
            <table style="border-collapse: collapse; table-layout: fixed; width: 100%;">
                <tr>
                    <th style="border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: center;">Nome da Impressora</th>
                    <th style="border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: center;">IP</th>
                    <th style="border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: center;">Horário</th>
                    <th style="border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: center;">Problema</th>
                </tr>
    """

    for printer in broke_printers:
        table += f"""
        <tr style="background-color: #ffffff">
            <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">{printer.get('Nome da Impressora')}</td>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">{printer.get('IP')}</td>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: left;">{printer.get('Horario')}</td>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: pre-wrap; font-family: Arial, sans-serif;"><pre>{printer.get('Problema')}</pre></td>
        </tr>
        """

    table += """
            </table>
        </body>
    </html>
    """

    return table

def sender_email(message_body):
    config = load_email_config()
    print("Enviando E-mail ...")
    try:
        email_smtp_server = config["server"]
        email_smtp_server_port = config["port"]
        email_smtp_user = config["user"]
        email_smtp_pass = config["pass"]

        msg = MIMEMultipart()
        msg['From'] = email_smtp_user
        msg['To'] = "ti@carmelhoteis.com.br"
        msg['Subject'] = f"Problema com impressora"
        body = email_body_formater(message_body)
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(email_smtp_server, email_smtp_server_port)
        server.starttls()
        server.login(email_smtp_user, email_smtp_pass)
        server.sendmail(email_smtp_user, "ti@carmelhoteis.com.br", msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")

def schedule_ping_for_printers(excel_list, broke_printers):
    for printer in excel_list:
        start_time = printer["Horario Inicial"]
        end_time = printer["Horario Final"]

        if isinstance(start_time, datetime):
            start_time = start_time.time()
        if isinstance(end_time, datetime):
            end_time = end_time.time()
        
        current_time = start_time
        printer_schedule = []
        while current_time <= end_time:
            time_str = current_time.strftime("%H:%M")
            schedule.every().day.at(time_str).do(ping_printers, printer=printer, broke_printers=broke_printers, time=time_str)
            print(f"Agendado para {printer['Nome da Impressora']} às {time_str}")

            current_time = (datetime.combine(datetime.today(), current_time) + timedelta(minutes=1)).time()

def get_session_token():
    session = requests.Session()
    session.auth = (auth.glpi_api.user.username, auth.glpi_api.user.password)
    response = session.get(auth.glpi_api.requests.newSession, headers={"Content-Type": "application/json", "App-Token": auth.glpi_api.appToken, "User-Token": auth.glpi_api.userToken})
    return response.json()["session_token"]

def create_new_ticket(body):
    session = requests.Session()
    session.auth = (auth.glpi_api.user.username, auth.glpi_api.user.password)
    token = get_session_token()
    respoonse = session.post(auth.glpi_api.requests.newTicket, headers={"Content-Type": "application/json", "App-Token": auth.glpi_api.appToken, "Session-Token": token}, json={
        "input": {
            "name": "Problema com a impressora testando o MAGNA",
            "content": html_content(body),
            "status": 2,
            "itilcategories_id": 128,
            "locations_id": 393,
            "entities_id": 1
        }
    })
    print(respoonse.status_code)
    print(respoonse.json())

def ping_printers(printer, broke_printers, time, num_packets=2):
    sent = 0
    received = 0
    lost = 0

    for _ in range(num_packets):
        packet = IP(dst=printer.get("IP")) / ICMP() / b"Ping personalizado!"
        response = sr1(packet, timeout=2, verbose=0)
        sent += 1
        if response:
            received += 1
        else:
            lost += 1

    if lost == 0:
        if printer in broke_printers:
            broke_printers.remove(printer)
    else:
        if printer not in broke_printers:
            printer["Horario"] = time
            printer["Problema"] = f"Estatísticas do Ping de {get_if_addr(conf.iface.name)}  para {printer.get("IP")}:\nPacotes: Enviados = {sent}, Recebidos = {received}, Perdidos = {lost}\n{(lost / sent) * 100}% de perda"
            broke_printers.append(printer)

    create_new_ticket(broke_printers)


# Cauãzinho gameplays falta só formatar para o modelo do email o corpo do chamado

