import cv2
from os import listdir
from os.path import join
import datetime as dt
import pytesseract
from pytesseract import Output


today = dt.datetime.now()

path = GetVar("workfolder_path")
messages = {messages}

folder = join(path, "reports", today.strftime("%d-%m-%Y"), "celar")

folder = join(path, "reports", today.strftime("%d-%m-%Y"), "celar")

files = [file for file in listdir(folder) if 'report' in file and file.endswith(".png")]

print(files)



for file, message in zip(files, messages):
    # Carga de la página 1 de novedad.pdf convertida a imagen png
    image = cv2.imread(join(folder, file))
    # image = imutils.resize(image, width=800, height=600)

    # Conversión de la imagen en escala de grises
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # acotar la región de interés de la imagen, la cual corresponde al código ventas
    gray_redux = gray_image[853:893 ,518:619]
    # plt.imshow(gray_image)
    # plt.show()

    # Umbralización de la zona de interés para realizar OCR
    # threshold_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    threshold_image = cv2.threshold(gray_redux, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

    # cv2.imshow("threshold_image", threshold_image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()

    # Obtener una cadena de texto al procesar la imagne umbralizada con PyTesseract
    custom_config = r'--oem 3 --psm 6'
    details = pytesseract.image_to_data(threshold_image, output_type=Output.DICT, config=custom_config, lang='spa')
    text = ''.join(details['text'])

    print(f"Code extracted: {text}")
    message["code"] = text


# SetVar("code", text)
SetVar("messages", messages)