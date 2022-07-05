import pandas as pd
import datetime as dt
from os.path import join, exists
from os import mkdir

path = GetVar("workfolder_path")
email_settings = {email_settings}

# Crear una variable con las zonas de las novedades
zones = email_settings['andina']['recipients']

# Definición de variable para la fecha de hoy.
today = dt.datetime.now()

# Definición variable ruta a reports
folder = join(path, 'reports')

# Definición de carpeta por fecha.
folder = join(folder, today.strftime("%d-%m-%Y"))

# Definición de carpeta para andinas.
folder = join(folder, 'andina')

# Crear carpeta novedades escaladas
scaled_folder = join(folder, 'escaladas')
if not exists(scaled_folder):
    mkdir(scaled_folder)


with open(join(folder, f'Novedades-{today.hour}h.xlsx'), mode='rb') as fp:
    df = pd.read_excel(fp, sheet_name='Novedades', engine='openpyxl')


for zone in zones:
    with pd.ExcelWriter(join(scaled_folder, f'Novedades {zone}-{today.hour}h.xlsx')) as writer:
        df_zone = df[df.frente == zone]
        # print(df_zone.head())
        df_zone.to_excel(writer, index=False, sheet_name=zone)

print("All records are classied by zone.")
