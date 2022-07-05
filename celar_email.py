import pandas as pd
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
email_subject = GetVar("email_subject")
msg = GetVar("email_message")
server = GetVar("smtp_server")
username = GetVar("smtp_username")
password = GetVar("smtp_password")
code = {code}
# Nombre de la base de datos
db_name = GetVar("db_name")
# Diccionario con los mensajes a enviar
messages = {messages}


# Nombre directorio empresa
enterprise_name = "celar"

# Nombre del archivo a adjuntar.
filename = "novedad.pdf"

# (L15-20) Definición de las variables para configuración del servicio de correo

# Declaración de una variable para la fecha de hoy
today = dt.datetime.now()

# Ruta de la base de datos
database_path = join(path, "database", db_name)

folder = join(path, "reports", today.strftime("%d-%m-%Y"), enterprise_name)

# Lectura de la base de datos y almacenamiento como dataframe
with open(database_path, mode='rb') as fp:
    df = pd.read_excel(fp, engine="openpyxl")

df.dropna(inplace=True, axis="columns")


for message in messages:
    df_filtered = df[(df['CODIGO SV'] == int(message.get('code')))]
    recipient = df_filtered['Correo Electrónico'].to_list()[0]
    sales_site = df_filtered['NOMBRE DEL SV'].to_list()[0]

    message["recipient"] = recipient
    message["site"] = sales_site


# 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje.
# 2. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo).
# 3. Configuración del asunto del correo (message['Subject']).
# 4. Adjuntar el cuerpo del correo al mensaje.
# 5. Definición de un objeto MIMEBase que sirve para cargar un archivo PDF y adjuntarlo en el mensaje.


for _msg in messages:
    # 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje

    # Configuración del formato del email
    message = MIMEMultipart()
    message['From'] = username

    # 2. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo)
    message['To'] = _msg.get('recipient')

    # 3. Configuración del asunto del correo (message['Subject'])
    subject = email_subject
    message["Subject"] = subject

    # 4. Adjuntar el cuerpo del correo al mensaje.
    # message.attach(MIMEText(msg, 'plain', 'utf8'))
    text = msg.replace('$(sitio)', _msg.get('site')).replace('$(codigo)', _msg.get('code'))
    # text = text.replace('$(codigo)', _msg.get('code'))
    message.attach(MIMEText(text, 'plain'))

    # 5. Definición de un objeto MIMEImage que sirve para cargar una imagen y adjuntarla en el mensaje.
    # ruta de la carpeta donde estan almacenado el archivo zip y el consolidado.
    folder = join(path, "reports", today.strftime("%d-%m-%Y"))
    report_path = join(folder, enterprise_name, _msg.get('report'))

    # Inserción del consolidado excel en el mensaje.
    with open(report_path, mode='rb') as part:
        pdf_file = MIMEBase('application', 'octet-stream')
        pdf_file.set_payload(part.read())

    encoders.encode_base64(pdf_file)
    pdf_file.add_header('Content-Disposition', 'attachment',
                        filename=filename)
    message.attach(pdf_file)

    # Envío de correo mediante una conexión SMTP a la bandeja de correo especificada en la configuración
    try:
        with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
            conn.starttls()
            conn.login(user=username, password=password)
            conn.sendmail(from_addr=username, to_addrs=['lmoreno@blacksmithresearch.com', 'jbustamante@blacksmithresearch.com'],
                        msg=message.as_string())
    except Exception as e:
        print(f"The connection has thrown an error: {e}")
        # file.write(
        #     f"ERROR [{dt.datetime.now()}] The connection with email server timed out.\n")
    else:
        print("The message has been sent successfully")
        # file.write(
        #     f"INFO [{dt.datetime.now()}] Report sucessfully sent to receivers_list.\n")
