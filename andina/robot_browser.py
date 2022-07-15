from logging import exception
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

# Se crea una bandera para evaluar el flujo del robot
flow_complete = 0
# Definición de variable para la fecha de hoy.
today = dt.datetime.today()

# Definición variable ruta a reports
folder = join(path, 'reports')

# Creación de carpeta por fecha.
folder = join(folder, today.strftime("%d-%m-%Y"))
if not exists(folder):
    mkdir(folder)

# Creación de carpeta para andinas.
# TODO Revisar la creación de esta carpeta
# folder = join(folder, 'andina')
# if not exists(folder):
    # mkdir(folder)

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
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
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
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
        return False
    else:
        return True


def statistics_page_available(driver: webdriver.Chrome) -> bool:
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "fecha")))
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
        return False
    else:
        return True


def projectFactsWrap_available(driver: webdriver.Chrome) -> bool:
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "projectFactsWrap")))
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
        return False
    else:
        return True

def chooseDate(driver: webdriver.Chrome, date: dt.datetime) -> bool:
    date_full = date.strftime("%Y-%m-%d")
    date_month = date.strftime("%B")
    date_year = date.strftime("%Y")
    try:
        # Buscar el botón para abrir el menú de selección para mes y año
        date_clicker = driver.find_element(By.CSS_SELECTOR, f"[aria-label='Choose month and year']")
        date_clicker.click()
        # Dar click sobre el botón del año
        date_clicker = driver.find_element(By.CSS_SELECTOR, f"[aria-label='{date_year}']")
        date_clicker.click()
        # Dar click sobre el botón del mes
        date_clicker = driver.find_element(By.CSS_SELECTOR,f"[aria-label='{date_month} {date_year}']")
        date_clicker.click()
        # Seleccionar la fecha requerida
        date_clicker = driver.find_element(By.CSS_SELECTOR,f"[aria-label='{date_full}']")
        date_clicker.click()
    except exception as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
        return False


def news(driver: webdriver.Chrome) -> bool:
    try:
        # Verificar si estan disponibles los divs con el selector css especificado
        divs = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.item.wow.fadeInUpBig.animated.animated")))
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )
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
    except TimeoutException as e:
        fp.write(
            f"Error [{dt.datetime.now()}] {e}.\n"
        )   
        return False
    else:
        return True

with open(join(folder, f"logAndina{today.strftime('%d-%m-%Y %Hh')}.txt"), mode='a', encoding='utf8') as fp:
    # (L254-257) Inicialización y configuración de objeto de la clase Chrome
    # Inicialización de un objeto de la clase Chrome para realizar la navegación en la WebSite de PowerBI.
    driver = webdriver.Chrome(
        executable_path=chrome_driver_path, chrome_options=chrome_options)
    driver.delete_all_cookies()

    # (L261-422) Ejecución del flujo de navegación del robot en Andina Seguridad.

    print(f"INFO [{dt.datetime.now()}] Robot execution started.")
    fp.write(
        f"INFO [{dt.datetime.now()}] Robot execution started.\n"
    )

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
            #TODO Introducir validación para el primer día del mes
            if statistics_page_available:
                # Calcular la fecha del dia anterior
                delta = dt.timedelta(days=21)
                yesterday = today - delta

                # Buscar los input de fecha inicio y fin y el btn de buscar.
                # Insertar fechas en los input para realizar la busqueda. 
                # (se modificaron los elementos a "readonly", por lo que se debe hacer click)
                # Se invoca el método chooseDate para elegir la fecha
                start_date = driver.find_element(By.ID, "fecha")
                start_date.click()
                #driver.execute_script(f'arguments[0].setAttribute("readonly", "false");',start_date)
                #start_date.send_keys('2022-07-14')
                chooseDate(driver, yesterday)

                end_date = driver.find_element(By.ID, "fecha_fin")
                end_date.click()
                chooseDate(driver, today)

                search_btn = driver.find_element(
                    By.CSS_SELECTOR, "button.btn.btn-primary")
                # Realizar la busqueda
                search_btn.click()

                if projectFactsWrap_available(driver):
                    # Verificar si existen novedades
                    news_list = []
                    if news(driver):
                        print("Hay novedades")
                        fp.write(
                            f"INFO [{dt.datetime.now()}] New findings.\n"
                        )
                        if app_available(driver):
                            # Extraer todas las tarjetas de novedades
                            # items = driver.find_elements(
                            #     By.CSS_SELECTOR, "div.card")
                            t.sleep(2)
                            items = WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.card")))
                            
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
                        fp.write(
                            f"INFO [{dt.datetime.now()}] Not new findings.\n"
                        )
                # Se levanta una bandera para confirmar el correcto flujo
                flow_complete = True

    print("Process finished.\n")
    fp.write(
            f"INFO [{dt.datetime.now()}] Process finished.\n"
    )
    print("Exiting...")
    fp.write(
            f"INFO [{dt.datetime.now()}] Exiting...\n"
    )

    SetVar("news", news_list)
    SetVar("flag_control", flow_complete)
    driver.quit()

