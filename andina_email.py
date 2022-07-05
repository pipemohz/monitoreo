from email import encoders
import smtplib
from os.path import join, dirname
import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

# Variables de entorno
path = GetVar('workfolder_path')
# Configuración del servicio de correo para enviar reportes por SMTP
port = {smtp_port}
email_settings = {email_settings}
msg = GetVar("email_message")
server = GetVar("smtp_server")
username = GetVar("smtp_username")
password = GetVar("smtp_password")

# Crear una variable con las zonas de las novedades
zones = email_settings['andina']['recipients']

# (L15-20) Definición de las variables para configuración del servicio de correo

# Declaración de una variable para la fecha de hoy
today = dt.datetime.now()

# 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje.
# 2. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo).
# 3. Configuración del asunto del correo (message['Subject']).
# 4. Adjuntar el cuerpo del correo al mensaje.
# 5. Definición de un objeto MIMEBase que sirve para cargar un archivo PDF y adjuntarlo en el mensaje.

# 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje

# Configuración del formato del email
message = MIMEMultipart()
message['From'] = username

# 2. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo)
message['To'] = ','.join(email_settings.get('recipients'))

# 3. Configuración del asunto del correo (message['Subject'])
subject = email_settings.get('subject').replace(
    '$(fecha)', today.strftime("%d-%m-%Y"))
message["Subject"] = subject

# 4. Adjuntar el cuerpo del correo al mensaje.
# message.attach(MIMEText(msg, 'plain', 'utf8'))
message.attach(MIMEText(msg, 'plain'))

# 5. Definición de un objeto MIMEImage que sirve para cargar una imagen y adjuntarla en el mensaje.
# ruta de la carpeta donde estan almacenado el archivo zip y el consolidado.
folder = join(path, "reports", today.strftime("%d-%m-%Y"))
report_path = join(folder, 'andina', f'Novedades{today.hour}h.xlsx')

# Inserción del consolidado excel en el mensaje.
with open(report_path, mode='rb') as part:
    excel_file = MIMEBase('application', 'octet-stream')
    excel_file.set_payload(part.read())

encoders.encode_base64(excel_file)
excel_file.add_header('Content-Disposition', 'attachment',
                      filename='Consolidado.xlsx')
message.attach(excel_file)

# Envío de correo mediante una conexión SMTP a la bandeja de correo especificada en la configuración
try:
    with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
        conn.starttls()
        conn.login(user=username, password=password)
        conn.sendmail(from_addr=username, to_addrs=email_settings.get('recipients'),
                      msg=message.as_string())
except Exception as e:
    print(f"The connection has thrown an error: {e}")
    # file.write(
    #     f"ERROR [{dt.datetime.now()}] The connection with email server timed out.\n")
else:
    print("The message has been sent successfully")
    # file.write(
    #     f"INFO [{dt.datetime.now()}] Report sucessfully sent to receivers_list.\n")
