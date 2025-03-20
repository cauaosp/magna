import smtplib
import subprocess
import openpyxl
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def ping_printers(printer):
    command = ["ping", printer.get("IP"), "-n", "2"]
    try:
        ping_result = subprocess.run(command, capture_output=True, text=True, encoding='latin-1')

        if ping_result.returncode == 0:
            if "Host de destino inacessível" in ping_result.stdout or "Host de destino inacess¡vel" in ping_result.stdout:
                return printer
        else:
            return printer
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


def read_sheet(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    header = [cell.value for cell in ws[1]]

    excel_list = []

    for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        linha_dict = {}
        for i, valor in enumerate(row):
            if valor is not None:
                linha_dict[header[i]] = valor
        if linha_dict.get("IP") is not None:
            excel_list.append(linha_dict)
    return excel_list


def load_email_config():
    # Carregando as configurações do arquivo auth.json
    with open('auth/auth.json', 'r') as f:
        data = json.load(f)
    
    email_config = data["email_smtp"]
    return email_config


def sender_email( message_body):
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
        body = f"""
        <html>
            <head>
                <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
            </head>
            <body>
                <p><strong>Problemas com as seguintes impressoras</strong>.</p>
                <p>{message_body}</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(email_smtp_server, email_smtp_server_port)
        server.starttls()
        server.login(email_smtp_user, email_smtp_pass)
        server.sendmail(email_smtp_user, "ti@carmelhoteis.com.br", msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")


