from genericpath import exists
from os import mkdir
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
# Nombre de la base de datos
db_name = GetVar("db_name")
# Diccionario con los mensajes a enviar
messages = {messages}
# Mensaje a enviar en caso de error
msg_error = "Los siguientes codigos no se encuentran en la base de datos:\n"
# Nombre directorio empresa
enterprise_name = "miro"


# (L15-20) Definición de las variables para configuración del servicio de correo

# Declaración de una variable para la fecha de hoy
today = dt.datetime.now()

# Ruta de la base de datos
database_path = join(path, "database", db_name)

folder = join(path, "reports", today.strftime("%d-%m-%Y"))
if not exists(folder):
    mkdir(folder)

folder = join(folder, enterprise_name)
if not exists(folder):
    mkdir(folder)

# Lectura de la base de datos y almacenamiento como dataframe
with open(database_path, mode='rb') as fp:
    df = pd.read_excel(fp, engine="openpyxl")

df.dropna(inplace=True, axis="columns")


for message in messages:
    df_filtered = df[(df['CODIGO SV'] == int(message.get('code')))]
    if not df_filtered.empty:
        recipient = df_filtered['Correo Electrónico'].to_list()[0]
        sales_site = df_filtered['NOMBRE DEL SV'].to_list()[0]
        print(message.get('code'))
        message["recipient"] = recipient
        message["site"] = sales_site

# 1. Busqueda y notificación de errores en el proceso de identificar los recipientes en la BD
# 2. Definición de un objecto MIMEMultipart para generar un nuevo mensaje.
# 3. Configuración de los recipientes para envío de correo (message['To'])
# 4. Configuración del asunto del correo (message['Subject']).
# 5. Adjuntar el cuerpo del correo al mensaje.

with open(join(folder, f"logMiro{today.strftime('%d-%m-%Y %Hh')}.txt"), mode='a', encoding='utf8') as fp:
    # 1. Busqueda y notificación de errores en el proceso de identificar los recipientes en la BD
    errors = [_msg for _msg in messages if _msg.get('recipient') == None]
    if errors:
        fp.write(
                f"Warning [{dt.datetime.now()}] Some sites don't have a code \
                in database ({len(errors)}/{len(messages)}).\n")
        
        for err in errors:
            msg_error += f"Code: {err.get('code')}\tLink: {err.get('href')}\n"
        print(msg_error)

        # 2. Definición de un objecto MIMEMultipart para generar un nuevo mensaje
        # Configuración del formato del email
        message = MIMEMultipart()
        message['From'] = username

        # 3. Configuración de los recipientes para envío de correo (message['To'])
        message['To'] = 'jbustamante@blacksmithresearch.com, ing.johneduardbt@gmail.com'

        # 4. Configuración del asunto del correo (message['Subject'])
        subject = f"Error en reporte de novedad {enterprise_name}"
        message["Subject"] = subject

        # 5. Adjuntar el cuerpo del correo al mensaje.
        text = msg_error + "\n\nEste correo se ha generado automáticamente, favor no responder este mensaje."
        message.attach(MIMEText(text, 'plain'))
        # Envío de correo mediante una conexión SMTP a la bandeja de correo especificada en la configuración
        try:
            with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
                conn.starttls()
                conn.login(user=username, password=password)
                conn.sendmail(from_addr=username, to_addrs=['jbustamante@blacksmithresearch.com'],
                            msg=message.as_string())
        except Exception as e:
            print(f"The connection has thrown an error: {e}")
            fp.write(
                f"ERROR [{dt.datetime.now()}] Error in sending error notification.\n")
        else:
            print("The error notification message has been sent successfully")
            fp.write(
                f"INFO [{dt.datetime.now()}] The error notification message has been sent successfully.\n")
    
    for _msg in messages:
        if _msg.get('recipient') is not None:
            # 2. Definición de un objecto MIMEMultipart para generar un nuevo mensaje

            # Configuración del formato del email
            message = MIMEMultipart()
            message['From'] = username

            # 3. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo)
            #message['To'] = _msg.get('recipient')

            # 4. Configuración del asunto del correo (message['Subject'])
            subject = email_subject
            message["Subject"] = subject

            # 5. Adjuntar el cuerpo del correo al mensaje.
            text = msg.replace('$(sitio)', _msg.get('site')).replace(
                '$(codigo)', _msg.get('code')).replace('$(enlace)', _msg.get('href'))
            message.attach(MIMEText(text, 'plain'))

            # Envío de correo mediante una conexión SMTP a la bandeja de correo especificada en la configuración
            try:
                with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
                    conn.starttls()
                    conn.login(user=username, password=password)
                    conn.sendmail(from_addr=username, to_addrs=['jbustamante@blacksmithresearch.com'],
                                msg=message.as_string())
            except Exception as e:
                print(f"The connection has thrown an error: {e}")
                fp.write(
                    f"ERROR [{dt.datetime.now()}] Error in sending notification of new in site {_msg['code']} to{_msg['recipient']}.\n")
            else:
                print("The message has been sent successfully")
                fp.write(
                    f"INFO [{dt.datetime.now()}] New in site {_msg['code']} notified to {_msg['recipient']}. Link to new: {_msg['href']}.\n")
