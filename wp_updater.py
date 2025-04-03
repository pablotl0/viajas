import os
import time
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options

# Configuración
EXCEL_FILE = r'backup\pablo-dominios.ods'
BACKUP_DIR = 'backup'
CHROME_PATH = r'C:\webdrivers\chrome-win64\chrome.exe'
CHROME_DRIVER_PATH = r'C:\webdrivers\chromedriver-win64\chromedriver.exe'
FILAS_A_PROCESAR = [12,17,19,20,22,23,24,25,26,27,28,29,30]  # Filas a procesar

def get_chrome_driver():
    service = Service(executable_path=CHROME_DRIVER_PATH)
    options = Options()
    options.binary_location = CHROME_PATH
    
    # Configuración para suprimir logs no deseados
    options.add_argument('--log-level=3')  # Solo errores críticos
    options.add_argument('--disable-logging')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # Configuración headless
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    
    # Deshabilitar servicios específicos que generan errores
    options.add_argument('--disable-notifications')
    options.add_argument('--disable-cloud-import')
    options.add_argument('--disable-gcm')
    options.add_argument('--disable-background-networking')
    
    # Configuración de descargas
    prefs = {
        "download.default_directory": os.path.abspath(BACKUP_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False  # Desactiva Safe Browsing para evitar logs
    }
    options.add_experimental_option("prefs", prefs)
    
    # Redirigir logs del driver
    service.log_path = 'NUL'  # Windows
    # service.log_path = '/dev/null'  # Linux/Mac
    
    return webdriver.Chrome(service=service, options=options)

def verificar_login_exitoso(driver):
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, 'wpadminbar'))
        )
        return True
    except TimeoutException:
        return False

def actualizar_wordpress(driver, dominio):
    try:
        driver.get(f"https://{dominio}/wp-admin/update-core.php")
        upgrade_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input#upgrade.button.button-primary.regular"))
        )
        upgrade_button.click()
        print("✓ WordPress actualizado")
        time.sleep(10)  # Esperar a que complete la actualización
        return True
    except TimeoutException:
        print("✓ WordPress ya está actualizado")
        return False
    except Exception as e:
        print(f"✗ Error al actualizar WordPress: {str(e)}")
        return False

def actualizar_plugins(driver, dominio):
    try:
        driver.get(f"https://{dominio}/wp-admin/update-core.php")
        # Seleccionar todos los plugins
        select_all = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input#plugins-select-all"))
        )
        select_all.click()
        
        # Hacer clic en actualizar
        upgrade_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input#upgrade-plugins.button"))
        )
        upgrade_button.click()
        print("✓ Plugins actualizados")
        time.sleep(10)  # Esperar a que complete la actualización
        return True
    except TimeoutException:
        print("✓ No hay plugins para actualizar")
        return False
    except Exception as e:
        print(f"✗ Error al actualizar plugins: {str(e)}")
        return False

def exportar_paginas(driver, dominio, backup_dir):
    try:
        # 1. Preparar directorio específico para el dominio
        dominio_sanitizado = dominio.replace('/', '_').replace(':', '_')
        domain_dir = os.path.abspath(os.path.join(backup_dir, dominio_sanitizado))
        os.makedirs(domain_dir, exist_ok=True)
        
        # 2. Configurar comportamiento de descarga
        params = {
            'behavior': 'allow',
            'downloadPath': domain_dir
        }
        driver.execute_cdp_cmd('Page.setDownloadBehavior', params)
        
        # 3. Navegar a la página de exportación
        driver.get(f"https://{dominio}/wp-admin/export.php")
        
        # 4. Seleccionar exportación de páginas
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='pages'][type='radio']"))
        ).click()
        
        # 5. Iniciar descarga
        submit_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input#submit.button.button-primary"))
        )
        submit_button.click()
        
        # 6. Esperar y verificar la descarga
        nombre_esperado = f"{dominio_sanitizado}_export.xml"
        ruta_esperada = os.path.join(domain_dir, nombre_esperado)
        
        # Patrón alternativo por si WordPress cambia el nombre
        patron_archivo = r"WordPress-\d{4}-\d{2}-\d{2}-\d{6}\.xml"
        
        tiempo_maximo = 30  # segundos
        intervalo = 2  # segundos
        tiempo_transcurrido = 0
        
        while tiempo_transcurrido < tiempo_maximo:
            time.sleep(intervalo)
            tiempo_transcurrido += intervalo
            
            # Buscar cualquier archivo XML reciente
            for archivo in os.listdir(domain_dir):
                if archivo.endswith('.xml'):
                    archivo_path = os.path.join(domain_dir, archivo)
                    # Renombrar si es necesario
                    if archivo != nombre_esperado:
                        try:
                            os.rename(archivo_path, ruta_esperada)
                            print(f"✓ Archivo renombrado a: {nombre_esperado}")
                        except Exception as rename_error:
                            print(f"⚠ No se pudo renombrar: {rename_error}")
                            ruta_esperada = archivo_path  # Usar el nombre original
                    
                    print(f"✓ Exportación exitosa: {ruta_esperada}")
                    return True
        
        print("✗ Tiempo de espera agotado - No se encontró el archivo descargado")
        return False
        
    except Exception as e:
        print(f"✗ Error durante la exportación: {str(e)}")
        return False

def leer_datos_ods(archivo_ods):
    """Lee un archivo ODS y devuelve los datos de las filas especificadas"""
    try:
        doc = load(archivo_ods)
        tabla = doc.getElementsByType(Table)[0]  # Asume que la primera tabla es la que contiene los datos
        
        datos = []
        for i, row in enumerate(tabla.getElementsByType(TableRow)):
            if i+1 in FILAS_A_PROCESAR:  # +1 porque las filas empiezan en 1
                celdas = row.getElementsByType(TableCell)
                if len(celdas) >= 4:  # Asegurarse de que hay al menos 4 columnas (A, B, C, D)
                    # Extraer dominio (columna A)
                    dominio = ''
                    if celdas[0].firstChild:
                        dominio = celdas[0].firstChild.data if hasattr(celdas[0].firstChild, 'data') else str(celdas[0].firstChild)
                    
                    # Extraer usuario (columna C)
                    usuario = ''
                    if len(celdas) > 2 and celdas[2].firstChild:
                        usuario = celdas[2].firstChild.data if hasattr(celdas[2].firstChild, 'data') else str(celdas[2].firstChild)
                    
                    # Extraer contraseña (columna D)
                    contraseña = ''
                    if len(celdas) > 3 and celdas[3].firstChild:
                        contraseña = celdas[3].firstChild.data if hasattr(celdas[3].firstChild, 'data') else str(celdas[3].firstChild)
                    
                    datos.append({
                        'fila': i+1,
                        'dominio': str(dominio).strip(),
                        'usuario': str(usuario).strip(),
                        'contraseña': str(contraseña).strip()
                    })
        return datos
    except Exception as e:
        print(f"Error al leer el archivo ODS: {str(e)}")
        return []

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"\nError: Archivo no encontrado en {EXCEL_FILE}")
        return

    datos = leer_datos_ods(EXCEL_FILE)
    if not datos:
        print("\nNo se encontraron datos válidos en el archivo ODS")
        return

    os.makedirs(BACKUP_DIR, exist_ok=True)

    for sitio in datos:
        dominio = sitio['dominio']
        usuario = sitio['usuario']
        contraseña = sitio['contraseña']
        fila = sitio['fila']

        if not all([dominio, usuario, contraseña]):
            print(f"\nFila {fila}: Datos incompletos")
            continue

        print(f"\n=== Procesando fila {fila}: {dominio} ===")
        driver = None
        try:
            # 1. Iniciar navegador y login
            driver = get_chrome_driver()
            driver.get(f"https://{dominio}/wp-admin/")
            
            driver.find_element(By.ID, 'user_login').send_keys(usuario)
            driver.find_element(By.ID, 'user_pass').send_keys(contraseña)
            driver.find_element(By.ID, 'wp-submit').click()
            
            if not verificar_login_exitoso(driver):
                print("✗ Falló el login")
                continue

            print("✓ Login exitoso")
                        
            # 2. Actualizar WordPress
            actualizar_wordpress(driver, dominio)
            
            # 3. Actualizar plugins
            actualizar_plugins(driver, dominio)
            
            # 4. Exportar páginas
            exportar_paginas(driver, dominio, BACKUP_DIR)

        except Exception as e:
            print(f"\n⚠️ Error inesperado: {str(e)}")
        finally:
            if driver:
                driver.quit()

    print("\n✅ Proceso finalizado")

if __name__ == "__main__":
    main()