import email
from os.path import join, exists
from os import mkdir
import imaplib
import datetime as dt

global mail_host, mail_username, mail_password, folder, connect, download_all_attachments, today

path = GetVar("workfolder_path")
mail_host = GetVar("smtp_server")
mail_username = GetVar("smtp_username")
mail_password = GetVar("smtp_password")

# Crear una variable para fecha-hora de hoy
today = dt.datetime.now()

# Crear una carpeta para reports
folder = join(path, "reports")

# Si no existe reports, crea el directorio
if not exists(folder):
    mkdir(folder)

folder = join(folder, today.strftime("%d-%m-%Y"))
# Si no existe la carpeta por fecha, crea el directorio
if not exists(folder):
    mkdir(folder)

# Crear una carpeta para celar
folder = join(folder, "celar")
# Si no existe la carpeta celar, crea el directorio
if not exists(folder):
    mkdir(folder)


def connect():
    """
    Create SSL connection to mailbox. Return a IMA4_SSL object.
    """

    conn = imaplib.IMAP4_SSL(mail_host)
    conn.login(mail_username,
               password=mail_password)
    conn.select(readonly=False)
    return conn


def download_all_attachments(conn: imaplib.IMAP4_SSL, email_id: bytes, index: int):
    """
    Download all files attached in message.
    """
    from os.path import join
    
    typ, data = conn.fetch(email_id, '(RFC822)')
    email_body = data[0][1]
    # print(email_body)
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            filename = part.get_filename().split('.')[0]
            print(filename)
            open(join(folder, f"{filename}{index + 1}.pdf"),
                 'wb').write(part.get_payload(decode=True))


def mailbox():
    """
    Stablish a connection with mailbox, search all messages with defined subject and download all files attached to message.
    """

    # conn = connect()

    with imaplib.IMAP4_SSL(mail_host) as conn:
        conn.login(mail_username,
               password=mail_password)
        conn.select(readonly=False)
        print(f'Fecha: {today.strftime("%d-%b-%Y")}')
        # Usar el metodo search para filtrar los mensajes por asunto y fecha. Necesario para adicionar el FLAG de Seen (Visto).
        data = conn.search(None, '(SUBJECT "RV: Reporte de novedad Celar Monitoreo SV")')[1]
        # data = conn.search(None, '(SUBJECT "RV: Reporte de novedad Celar Monitoreo SV"', f'ON {today.strftime("%d-%b-%Y")}', 'UNSEEN)')[1]
        # data = conn.uid('search', '(SUBJECT "RV: Reporte de novedad Celar Monitoreo SV"', f'ON 30-Jun-2022', 'UNSEEN)')[1]
        emails_id = [_id.decode('utf8') for _id in data[0].split()] 
        print(emails_id)

        print(f"Mensajes encontrados: {len(emails_id)}")
        if emails_id:
            for index, _id in enumerate(emails_id):
                download_all_attachments(conn, _id, index)
                conn.store(_id, '+FLAGS', r'(\Seen)')
            # conn.expunge()

    # #Crear un nuevo mailbox
    # response = conn.create('Celar')
    # print(f"response{response}")

    # Mover mensajes a nuevo mailbox
    # conn.copy(emails_id[0].decode(), 'Celar')
    # conn.close()
    # conn.logout()


mailbox()
