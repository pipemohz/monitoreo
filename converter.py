from pdf2image import convert_from_path
from os.path import join
from os import listdir
import datetime as dt

path = GetVar("workfolder_path")
filename = "novedad.pdf"

# Declaración variable fecha
today = dt.datetime.now()

folder = join(path, "reports", today.strftime("%d-%m-%Y"), "celar")

files = [file for file in listdir(folder) if 'novedad' in file]

# Lista de diccionarios de mensajes con codigo y reporte como keys asociadas.
messages = []

for index, file in enumerate(files):
    convert_from_path(join(folder, file), output_folder=folder,
                      output_file=f"report{index + 1}", fmt="png", single_file=True)

    messages.append({"report": file})
    print(f"File {file} converted.")

print("All pdf files converted to png")

SetVar("messages", messages)
