import json
from os import mkdir
from os.path import join, exists

path = GetVar("workfolder_path")
enterprise_name = "miro"

with open(join(path, "config.json"), mode='r', encoding='utf8') as fp:
    data = json.load(fp)

# Crear la carpeta de reportes.
folder = join(path, "reports")
# Si la carpeta 'reports' no existe la crea.
if not exists(folder):
    mkdir(folder)

# Crear la carpeta de templates.
folder = join(path, "templates")
if not exists(folder):
    mkdir(folder)

# Crear la carpeta de base de datos.
folder = join(path, "database")
if not exists(folder):
    mkdir(folder)

# Crear la carpeta de email.
folder = join(path, "email")
if not exists(folder):
    mkdir(folder)

with open(join(folder, data['email']['email_settings']), encoding='utf8') as fp:
    email_settings = json.load(fp)
    email_subject = email_settings[enterprise_name]['subject']
    email_message = email_settings[enterprise_name]['message']

with open(join(folder, email_message), encoding='utf8') as file:
    msg = ""
    for line in file.readlines():
        msg += line

SetVar("smtp_server", data['email']['smtp_server'])
SetVar("smtp_port", data['email']['smtp_port'])
SetVar("smtp_username", data['email']['smtp_username'])
SetVar("smtp_password", data['email']['smtp_password'])
SetVar("email_subject", email_subject)
SetVar("email_message", msg)
SetVar("db_name", data.get("db_name"))
