from email import encoders
import smtplib
from os.path import join, dirname
import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

from numpy import empty

# Variables de entorno
path = GetVar('workfolder_path')
# Configuración del servicio de correo para enviar reportes por SMTP
port = {smtp_port}
email_subject = "{email_subject}"
msg = GetVar("email_message")
server = GetVar("smtp_server")
username = GetVar("smtp_username")
password = GetVar("smtp_password")
news_list = {news}

# Nombre directorio empresa
enterprise_name = "Andina"
# Declaración de una variable para la fecha de hoy
today = dt.datetime.now()
# ruta de la carpeta donde estan almacenado el archivo zip y el consolidado.
folder = join(path, "reports", today.strftime("%d-%m-%Y"))
#report_path = join(folder, 'andina', f'Novedades{today.hour}h.xlsx')

with open(join(folder, f"logAndina{today.strftime('%d-%m-%Y %Hh')}.txt"), mode='a', encoding='utf8') as fp:
    fp.write(
        f"INFO [{dt.datetime.now()}] Start Email Service Robot.\n"
    )
    # Crear una variable con las zonas de las novedades
    #zones = email_settings['andina']['recipients']
    if len(news_list):
        for _new in news_list:
            # Se genera un registro en el archivo log de las novedades encontradas
            fp.write(
                        f"INFO [{dt.datetime.now()}] New in site {_new['office']}. Description: {_new['description']}. Link to image: {_new['src']}.\n")
        # (L15-20) Definición de las variables para configuración del servicio de correo

        # 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje.
        # 2. Configuración de los recipientes para envío de correo (message['To']) (Pendiente lista de correo).
        # 3. Configuración del asunto del correo (message['Subject']).
        # 4. Adjuntar el cuerpo del correo al mensaje.
        # 5. Definición de un objeto MIMEBase que sirve para cargar un archivo PDF y adjuntarlo en el mensaje.

        # 1. Definición de un objecto MIMEMultipart para generar un nuevo mensaje

        # Configuración del formato del email
            message = MIMEMultipart()
            message['From'] = username

            # 2. Configuración de los recipientes para envío de correo (message['To']) 
            # TODO (Pendiente lista de correo)
            #message['To'] = ','.join()

            # 3. Configuración del asunto del correo (message['Subject'])
            subject = email_subject.replace(
                '$(fecha)', today.strftime("%d-%m-%Y"))
            message["Subject"] = subject

            # 4. Adjuntar el cuerpo del correo al mensaje.
            # message.attach(MIMEText(msg, 'plain', 'utf8'))
            
            body = msg.replace('$(oficina)', _new.get('office'))\
                    .replace('$(novedad)', _new.get('description'))\
                    .replace('$(link)', _new.get('src'))
            message.attach(MIMEText(body, 'plain'))
            # Envío de correo mediante una conexión SMTP a la bandeja de correo especificada en la configuración
            try:
                with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
                    conn.starttls()
                    conn.login(user=username, password=password)
                    conn.sendmail(from_addr=username, to_addrs='jbustamante@blacksmithresearch.com',
                                msg=message.as_string())
            except Exception as e:
                print(f"The connection has thrown an error: {e}")
                fp.write(
                    f"ERROR [{dt.datetime.now()}] The connection with email server timed out.\n")
            else:
                print("The message has been sent successfully")
                fp.write(
                        f"INFO [{dt.datetime.now()}] Report sucessfully sent.\n")
    else:
        print("Not news to send")
        fp.write(
                f"INFO [{dt.datetime.now()}] Not news to send.\n")
    print("Process finished.\n")
    fp.write(
            f"INFO [{dt.datetime.now()}] Process finished.\n")