import email
from email.errors import HeaderDefect
from os.path import join, exists
from os import mkdir
import imaplib
import datetime as dt

global mail_host, mail_username, mail_password, folder, connect, get_email_text, today

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


def get_email_text(conn: imaplib.IMAP4_SSL, email_id: bytes) -> dict:
    """
    Download all files attached in message.
    """

    # typ, data = conn.fetch(email_id, '(RFC822)')
    typ, data = conn.fetch(email_id, '(UID BODY[TEXT])')

    email_body = data[0][1]

    msg = email_body[email_body.find(b"Puesto Visitado"):].split(b'\r')[0]
    code = msg.split(b' ')[-1].decode('utf8')

    msg = email_body[email_body.find(b"https:"):].split(b' ')[0]
    href = msg.replace(b"\n", b"").replace(
        b"=\r", b"").replace(b"3D", b"").decode("utf8")

    return {"code": code, "href": href}


def mailbox() -> list:
    """
    Stablish a connection with mailbox, search all messages with defined subject and download all files attached to message.
    """

    with imaplib.IMAP4_SSL(mail_host) as conn:
        conn.login(mail_username,
                   password=mail_password)
        conn.select(readonly=False)
        # Usar el metodo search para filtrar los mensajes por asunto y fecha. Necesario para adicionar el FLAG de Seen (Visto).
        data = conn.search(None, '(SUBJECT "NOVEDAD REPORTES DE VISITA - REDITOS EMPRESARIALES Version 2.0"',
                           f'ON {today.strftime("%d-%b-%Y")}', 'UNSEEN)')[1]
        emails_id = [_id.decode('utf8') for _id in data[0].split()]
        # print(emails_id)

        messages = []
        print(f"Mensajes encontrados: {len(emails_id)}")
        if emails_id:
            for _id in emails_id:
                message = get_email_text(conn, _id)
                # get_email_text(conn, _id)
                messages.append(message)

                # conn.store(_id, '+FLAGS', r'(\Seen)')

    return messages


messages = mailbox()
print("All codes extracted from email messages.")
SetVar("messages", messages)
