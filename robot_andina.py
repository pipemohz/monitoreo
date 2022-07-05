from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import datetime as dt
import time as t
from os import mkdir, remove
from os.path import join, isdir, exists


# Obtener variables de proyecto
path = GetVar("workfolder_path")
chrome_driver_path = GetVar("chrome_driver_path")
credentials = {andina}

# Configuración de variables de control para el robot.

# Nombre del reporte de novedades descargado.
report_name = "Novedades.xlsx"

# Tiempo maximo de espera en la descarga del reporte
max_time = 10

# Definición de variable para la fecha de hoy.
today = dt.datetime.today()

# Definición variable ruta a reports
folder = join(path, 'reports')

# Creación de carpeta por fecha.
folder = join(folder, today.strftime("%d-%m-%Y"))
if not exists(folder):
    mkdir(folder)

# Creación de carpeta para andinas.
folder = join(folder, 'andina')
if not exists(folder):
    mkdir(folder)

# (L47-49) Declaración de un objeto ChromeOptions para configurar la ruta de descargas del navegador.
# Opciones del navegador
chrome_options = webdriver.ChromeOptions()
prefs = {"download.default_directory": folder}
chrome_options.add_experimental_option("prefs", prefs)

# (L58-248) Declaración de funciones para controlar el flujo de la sesión de navegación del robot.


def web_available(driver: webdriver.Chrome) -> bool:
    """
    Espera por la carga de la página principal de Andina Seguridad. 
    Retorna True si la página carga correctamente.
    False si la página no esta disponible antes del timeout.
    """
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/app-root/app-login/div/div/div/div/div/form/div[2]/span[1]/div/input")))
    except TimeoutException:
        return False
    else:
        return True


def user_panel_available(driver: webdriver.Chrome) -> bool:
    """
    Espera por la carga del panel de usuario de Andina Seguridad. 
    Retorna True si la página carga correctamente.
    False si la página no esta disponible antes del timeout.
    """
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "card-title")))
    except TimeoutException:
        return False
    else:
        return True


def statistics_page_available(driver: webdriver.Chrome) -> bool:
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "fecha")))
    except TimeoutException:
        return False
    else:
        return True


def projectFactsWrap_available(driver: webdriver.Chrome) -> bool:
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "projectFactsWrap")))
    except TimeoutException:
        return False
    else:
        return True


def news(driver: webdriver.Chrome) -> bool:
    try:
        # Verificar si estan disponibles los divs con el selector css especificado
        divs = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.item.wow.fadeInUpBig.animated.animated")))
    except TimeoutException:
        return False
    else:
        # Seleccionar el div 3 de la lista divs
        div = divs[2]
        number = int(div.find_element(
            By.TAG_NAME, "p#number2.number").get_attribute('innerText'))
        img = div.find_element(By.CSS_SELECTOR, "img")

        if number:
            driver.execute_script('arguments[0].click();', img)
            return True
        else:
            return False


def app_available(driver: webdriver.Chrome) -> bool:
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "app-novedades-qr")))
    except TimeoutException:
        return False
    else:
        return True


# (L254-257) Inicialización y configuración de objeto de la clase Chrome
# Inicialización de un objeto de la clase Chrome para realizar la navegación en la WebSite de PowerBI.
driver = webdriver.Chrome(
    executable_path=chrome_driver_path, chrome_options=chrome_options)
driver.delete_all_cookies()

# (L261-422) Ejecución del flujo de navegación del robot en Andina Seguridad.

print(f"INFO [{dt.datetime.now()}] Robot execution started.")

driver.get(url=credentials.get('URL'))

if web_available(driver):
    # Buscar el input de usuario en el form de inicio de sesión de Andina.
    # user_input = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div/div/div/form/div[2]/span[1]/div/input")
    user_input = driver.find_element(By.NAME, "email")
    # Insertar username
    user_input.send_keys(credentials.get('username'))

    # Buscar el input de password en el form de inicio de sesión de Andina.
    # pass_input = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div/div/div/form/div[2]/span[2]/div/input")
    pass_input = driver.find_element(By.NAME, "password")
    # Insertar password
    pass_input.send_keys(credentials.get('password'))

    # Buscar el btn de submit en el form de inicio de sesión de Andina.
    submit_btn = driver.find_element(
        By.XPATH, "/html/body/app-root/app-login/div/div/div/div/div/form/div[3]/div/div[1]/button")
    submit_btn.click()

    if user_panel_available(driver):
        statistics_card = driver.find_element(By.CLASS_NAME, "card-title")
        statistics_card.click()

        if statistics_page_available:
            # Calcular la fecha del dia anterior
            delta = dt.timedelta(days=21)
            yesterday = today - delta

            # Buscar los input de fecha inicio y fin y el btn de buscar.
            start_date = driver.find_element(By.ID, "fecha")
            end_date = driver.find_element(By.ID, "fecha_fin")
            search_btn = driver.find_element(
                By.CSS_SELECTOR, "button.btn.btn-danger")

            # Insertar fechas en los input para realizar la busqueda
            start_date.send_keys(yesterday.strftime("%d-%m-%Y"))
            end_date.send_keys(today.strftime("%d-%m-%Y"))
            # Realizar la busqueda
            search_btn.click()

            if projectFactsWrap_available(driver):
                # Verificar si existen novedades
                if news(driver):
                    print("Hay novedades")
                    if app_available(driver):
                        # Extraer todas las tarjetas de novedades
                        # items = driver.find_elements(
                        #     By.CSS_SELECTOR, "div.card")
                        t.sleep(2)
                        items = WebDriverWait(driver, 10).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.card")))
                        news_list = []
                        for item in items[1:]:
                            # Extraer elementos html de la card de novedades.
                            p_list = item.find_elements(By.CSS_SELECTOR, "p")
                            img = item.find_element(By.CSS_SELECTOR, "img")

                            src = img.get_attribute('src')
                            office = p_list[0].text
                            description = p_list[4].text
                            news_list.append(
                                {"office": office, "description": description, "src": src})

                else:
                    print("No hay novedades")

print("Process finished.")
print("Exiting...")

SetVar("news", news_list)
driver.quit()
