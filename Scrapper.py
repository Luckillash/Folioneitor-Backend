import ipaddress
import os
import sys
import time
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from SQL import main

if len(sys.argv) < 2 or len(sys.argv) > 3:

    print(f"Solo se aceptan 2 argumentos.")

    sys.exit(1)

try:

    ipaddress.ip_address(sys.argv[1])

    print(f"{sys.argv[1]} es una dirección IP válida.")

except:

    print(f"{sys.argv[1]} no es una dirección IP válida.")

    sys.exit(1)

cert_path = "C:\\Users\\lsalazar\\OneDrive - VM\\Escritorio\\Documentación de sistemas\\10042060-0.pfx"

cert_password = 'IvmCvm876'

# Directorio de descarga
download_dir = "C:\\CAF"

# Asegúrate de que el directorio de descarga exista
if not os.path.exists(download_dir):
    
    os.makedirs(download_dir)

options = Options()

options.add_argument('--ignore-certificate-errors')

options.add_argument('--allow-running-insecure-content')

options.add_argument('--no-sandbox')

# Configuración del certificado
options.add_argument(f'--ssl-client-certificate-file={cert_path}')

options.add_argument(f'--ssl-client-certificate-password={cert_password}')

options.add_argument("--window-size=1920,1080")

# Configuración de preferencias para descargas
prefs = {

    "download.default_directory": download_dir,

    "download.prompt_for_download": False,

    "directory_upgrade": True,

    "safebrowsing.enabled": True

}

options.add_experimental_option("prefs", prefs)

# Configuración del driver de Chrome
driver = webdriver.Chrome(options)

try:

    driver.get('https://palena.sii.cl/cvc_cgi/dte/of_solicita_folios')

    time.sleep(4)

    try:

        alert = driver.switch_to.alert

        alert.accept()

    except Exception as e:

        print(f"No se encontró ninguna alerta. {e}")

    time.sleep(2)

    x, y = 900, 700

    pyautogui.moveTo(x, y)

    time.sleep(5)

    pyautogui.click()

    time.sleep(4)

    try:

        rut_emp_input = driver.find_element(By.NAME, 'RUT_EMP')

        rut_emp_input.clear()

        rut_emp_input.send_keys('76126876')

        dv_emp_input = driver.find_element(By.NAME, 'DV_EMP')

        dv_emp_input.clear()

        dv_emp_input.send_keys('7')

        aceptar_button = driver.find_element(By.NAME, 'ACEPTAR')

        aceptar_button.click()

    except Exception as e:

        print(f"Error al encontrar un elemento: {e}")

    time.sleep(5)

    try:

        select_element = driver.find_element(By.NAME, 'COD_DOCTO')

        select = Select(select_element)

        select.select_by_value('39') 

        time.sleep(2)

        cant_doctos_input = driver.find_element(By.NAME, 'CANT_DOCTOS')

        cant_doctos_input.clear()

        cantidad_folios = '30000'

        if sys.argv[1] == '192.168.1.54':

            cantidad_folios = '100000'
            
        elif sys.argv[1] == '192.168.1.55':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.1.56':

            cantidad_folios = '10000'

        elif sys.argv[1] == '192.168.1.57':

            cantidad_folios = '25000'

        elif sys.argv[1] == '192.168.1.58':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.1.59':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.10.78':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.10.84':

            cantidad_folios = '40000'

        elif sys.argv[1] == '192.168.10.85':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.10.86':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.10.87':

            cantidad_folios = '30000'

        elif sys.argv[1] == '192.168.10.88':

            cantidad_folios = '20000'

        cant_doctos_input.send_keys(cantidad_folios)

        time.sleep(2)

        aceptar_button = driver.find_element(By.NAME, 'ACEPTAR')

        aceptar_button.click()

        time.sleep(2)

        acepta2_button = driver.find_element(By.NAME, 'ACEPTAR')

        acepta2_button.click()

        time.sleep(2)

        # Obtén una lista de archivos en el directorio de descargas antes de la descarga
        before_download = os.listdir(download_dir)

        generar_folios_button = driver.find_element(By.NAME, 'ACEPTAR')

        generar_folios_button.click()

        # Espera un momento para que se complete la descarga
        time.sleep(10)  # Ajusta este tiempo según sea necesario

        # Obtén una lista de archivos en el directorio de descargas después de la descarga
        after_download = os.listdir(download_dir)

        # Identifica el nuevo archivo descargado
        new_files = list(set(after_download) - set(before_download))

        if new_files:

            downloaded_file = new_files[0]

            print(f"Archivo descargado: {downloaded_file}")

            main(downloaded_file, sys.argv[1], sys.argv[2])

        else:

            print("No se encontró ningún archivo nuevo.")

    except Exception as e:

        print(f"Error al encontrar un elemento: {e}")

    time.sleep(3)

    print("NEXT")

finally:
    
    driver.quit()
