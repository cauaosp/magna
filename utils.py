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

def ping_printers(printer, broke_printers, time, num_packets=2):
    command = ["ping", printer.get("IP"), "-n", "2"]
    
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
    
    sender_email(broke_printers)


    # try:
    #     ping_result = subprocess.run(command, capture_output=True, text=True, encoding='latin-1')

    #     if ping_result.returncode == 0:
    #         if "Host de destino inacessível" in ping_result.stdout or "Host de destino inacess¡vel" in ping_result.stdout:
    #             printer["Problema"] = ping_result.stdout
    #             if printer not in broke_printers:
    #                 broke_printers.append(printer)
    #         else:
    #             if printer in broke_printers:
    #                 broke_printers.remove(printer)
    #     else:
    #         printer["Problema"] = ping_result.stdout
    #         if printer not in broke_printers:
    #             broke_printers.append(printer)

    #     printer["Horario"] = time
    #     print(broke_printers)
    #     # sender_email(broke_printers)
    # except Exception as e:
    #     print(f"Ocorreu um erro: {e}")
    #     return None

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

def email_body_formater(broke_printers):
    table = """
    <html>
        <head>
            <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
            <style>
                table {
                    border-collapse: collapse;
                    table-layout: fixed;
                }
                th, td {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                th {
                    background-color: #f2f2f2;
                    text-align: center;
                }
                pre {
                    white-space: pre-wrap; /* Preserva a formatação do texto */
                    font-family: Arial, sans-serif;
                }
            </style>
        </head>
        <body>
            <p><strong>Problemas com as seguintes impressoras:</strong></p>
            <table>
                <tr>
                    <th>Nome da Impressora</th>
                    <th>IP</th>
                    <th>Horário</th>
                    <th>Problema</th>
                </tr>
    """

    for printer in broke_printers:
        table += f"""
        <tr>
            <td>{printer.get('Nome da Impressora')}</td>
            <td>{printer.get('IP')}</td>
            <td>{printer.get('Horario')}</td>
            <td><pre>{printer.get('Problema')}</pre></td>
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
    response = requests.get(auth.glpi_api.requests.newSession, headers={
        "Content-Type": "application/json",
        "App-Token": auth.glpi_api.appToken,
        "User-Token": auth.glpi_api.userToken
    })
    print(response.json())
    print(auth)