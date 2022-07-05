import datetime as dt
from shutil import move
from os.path import join
from os import listdir

path = GetVar("workfolder_path")

# Nombre del reporte de novedades descargado.
report_name = "novedad"

# Definición de variable para la fecha de hoy.
today = dt.datetime.now()

# Definición variable ruta a reports
folder = join(path, 'reports')

# Definición de carpeta por fecha.
folder = join(folder, today.strftime("%d-%m-%Y"))

# Definición de carpeta para andinas.
folder = join(folder, 'celar')

reports = [file for file in listdir(folder) if "novedad" in file]

for index, report in enumerate(reports):
    move(join(folder, report), join(folder, f'novedad{index + 1}-{today.hour}h.pdf'))


print("File renamed successfully.")
