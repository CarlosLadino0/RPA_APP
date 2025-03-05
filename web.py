from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from datetime import datetime
from pywinauto.keyboard import send_keys
import unicodedata
import requests
import random
import time
import json
import os

API_KEY = "014c69e51a56b400e1d97e9c7a8d5f85"

def resolver_captcha(driver):
    print("Esperando a que el CAPTCHA se cargue...")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "g-recaptcha"))
    )

    print("Resolviendo reCAPTCHA v2 con 2Captcha")
    url_2captcha = "https://2captcha.com/in.php"
    params = {
        "key": API_KEY,
        "method": "userrecaptcha",
        "googlekey": "6LdJbMIqAAAAAF8OX8AxNy-TMJwPBwFL8Ty2ZEq6",
        "pageurl": "https://crm.agentemotor.com/avs/login?tenant=corredoresasociados.co.agentemotor.com",
        "json": 1
    }
    try:
        response = requests.get(url_2captcha, params=params).json()
    except Exception as e:
        print(f"Error al conectar con 2Captcha: {e}")
        return None
    
    if response["status"] != 1:
        print("Error al solicitar la resolución del CAPTCHA", response)
        return None

    request_id = response["request"]
    print(f"Captcha enviado, ID: {request_id}. Esperando respuesta...")

    url_result = f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={request_id}&json=1"
    
    for _ in  range(30):
        time.sleep(5)
        resultado = requests.get(url_result).json()
        if resultado["status"] == 1:
            print("Captcha resuelto exitosamente")
            return resultado["request"]
    
    print("Error al resolver el CAPTCHA")
    return None

driver = webdriver.Chrome()
driver.get("https://crm.agentemotor.com/avs/login?tenant=corredoresasociados.co.agentemotor.com")
driver.maximize_window()

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "login"))).send_keys("practicante@correseguros.co")
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys("Carlos2024*")

token_resuelto = resolver_captcha(driver)
if token_resuelto:
    driver.execute_script(f'document.getElementById("g-recaptcha-response").value="{token_resuelto}";')
    driver.execute_script("document.getElementById('g-recaptcha-response').dispatchEvent(new Event('change'));")
    time.sleep(2)

    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    print("Inicio de sesión enviado con el CAPTCHA resuelto.")
else:
    print("No se pudo resolver el CAPTCHA.")

def normalizar_texto(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto) 
        if unicodedata.category(c) != 'Mn'
    ).upper().strip()

def borrar_y_escribir(campo, texto_a_escribir):
        campo.click()
        texto_actual = campo.get_attribute('value')
        if texto_actual:
            campo.send_keys(Keys.BACKSPACE * len(texto_actual))
        campo.send_keys(texto_a_escribir)

def agente_motor():
    try:
        ruta_json = "datos/datos_extraidos.json"
        if not os.path.exists(ruta_json):
            raise FileNotFoundError(f"No se encontraron los datos extraídos en {ruta_json}")
        with open(ruta_json, "r", encoding="utf-8") as archivo_json:
            datos = json.load(archivo_json)

        clase_vehiculo = normalizar_texto(datos["datos_tarjeta"]["clase_vehiculo"])
        servicio = normalizar_texto(datos["datos_tarjeta"]["servicio"])
        driver.set_page_load_timeout(20)
        time.sleep(5)

        usuario = "practicante@correseguros.co"
        campo_usuario = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "login"))
            )
        campo_usuario.send_keys(usuario)

        time.sleep(1)

        '''
        for char in usuario:
            campo_usuario.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        '''        

        contrasena = "Carlos2024*"
        campo_contrasena = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        campo_contrasena.send_keys(contrasena)

        '''
        for char in contrasena:
            campo_contrasena.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))
        '''

        time.sleep(0.5)
        campo_contrasena.send_keys(Keys.TAB)
        time.sleep(0.5)     

        driver.switch_to.active_element.send_keys(Keys.TAB)
        time.sleep(0.5)
        driver.switch_to.active_element.send_keys(Keys.ENTER)

        boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
        boton_login.click()
        time.sleep(2)

        boton_cotizar = driver.find_element(By.XPATH, "//a[@href='/web#menu_id=280&action=']")
        boton_cotizar.click()
        time.sleep(3)

        opciones_botones = {
            "AUTOMÓVIL": {
                "PARTICULAR": "Particulares",
                "PÚBLICO": "Particulares"
            },
            "MOTOCICLETA": {
                "PARTICULAR": "Motos",
                "PUBLICO": "Motos"
            },
            "BUS":  {
                "PÚBLICO": "Públicos",
            },
            "BUSETA": {
                "PÚBLICO": "Públicos",
            },
            "CAMIÓN": {
                "PARTICULAR": "Carga",
                "PÚBLICO": "Carga",
            },
            "CAMIONETA": {
                "PARTICULAR": "Particulares",
                "PÚBLICO": "Públicos",
            },
            "MICROBUS": {
                "PÚBLICO": "Públicos",
            }
        }

        opciones_botones = {
            normalizar_texto(k): {normalizar_texto(sk): sv for sk, sv in v.items()} for k, v in opciones_botones.items()
            }
        if clase_vehiculo not in opciones_botones:
            print(f"No se maneja la clase de vehículo: {clase_vehiculo}")
            return
        if servicio not in opciones_botones[clase_vehiculo]:
            print(f"No se maneja el servicio: {servicio} para {clase_vehiculo}")
            return

        boton = opciones_botones[clase_vehiculo][servicio]

        try:
            num_cedula = datos["datos_cedula"]["num_cedula"]
            placa = datos["datos_tarjeta"]["placa"]
            nuevo_vehiculo = datos["nuevo_vehiculo"].strip().upper()
            modelo_esperado = datos["datos_tarjeta"]["modelo"]
            ciudad_circulacion = datos["zona_circulacion"]["ciudad"]
            nombres = datos["datos_cedula"]["nombres"]
            apellidos = datos["datos_cedula"]["apellidos"]
            fecha_nacimiento_raw = datos["datos_cedula"]["fecha_nacimiento"]
            genero = datos["genero"]
            oneroso = datos["oneroso"]
            marca = datos["datos_tarjeta"]["marca"]
            linea = datos["datos_tarjeta"]["linea"]
            correo = datos["correo"]

            fecha_nacimiento = datetime.strptime(fecha_nacimiento_raw, "%d-%b-%Y").strftime("%d-%m-%Y")
            print(f"Fecha de nacimiento convertida: {fecha_nacimiento}")

            selector = f'img[alt="{boton}"]'
            print(f"Buscando el elemento con selector: {selector}")

            elemento = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
            )
            elemento.click()
            print(f"Seleccionado: {clase_vehiculo} - {servicio} -> {boton}")
        except Exception as e:
            print(f"No se pudo seleccionar el botón ({clase_vehiculo} - {servicio}): {e}")
        time.sleep(2)

        try: 
            campo_cedula = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "identification_number"))
            )
            campo_cedula.send_keys(num_cedula)

            campo_placa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "plate"))
            )
            campo_placa.send_keys(placa)
            print(f"Se ingresaron los datos de la placa - {placa}. Y cédula - {num_cedula}.")

            wait = WebDriverWait(driver, 1)

            if nuevo_vehiculo == "SI":
                boton_si = wait.until(
                    EC.element_to_be_clickable((By.ID, "in_agency"))
                )
                boton_si.click()
                print("Se seleccionó 'Vehículo nuevo'")

                campo_modelo = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "model"))
                )
                campo_modelo.send_keys(modelo_esperado)
                campo_modelo.send_keys(Keys.ARROW_DOWN)
                campo_modelo.send_keys(Keys.ENTER)
                print(f"Se ingresó el modelo del vehículo: {modelo_esperado}")

                campo_marca = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "brand"))
                )
                time.sleep(1)
                campo_marca.send_keys(marca)
                time.sleep(1)
                campo_marca.send_keys(Keys.ARROW_DOWN)
                campo_marca.send_keys(Keys.ENTER)
                print(f"Se ingresó la marca del vehículo: {marca}")

                campo_linea = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "line"))
                )
                campo_linea.send_keys(linea)
                campo_linea.send_keys(Keys.ENTER)
                print(f"Se ingresó la línea del vehículo: {linea}")

            elif nuevo_vehiculo == "NO":
                print("No se pulsará 'Vehículo nuevo'")
            else:
                print(f"Valor inesperado para 'nuevo_vehiculo': {nuevo_vehiculo}")

        except Exception as e:
            print(f"No se pudieron llenar los campos placa - {placa}. Y cédula - {num_cedula}. {e}")

        try: 
            wait = WebDriverWait(driver, 5)
            boton_siguiente2 = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button-stepper-next"]'))
                )
            boton_siguiente2.click()
        except Exception as e:
            print(f"Error al pulsar el botón siguiente: {e}")

        try:
            mensaje_no_encontrado = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//h4[text()='Vehículo no encontrado!!']"))
            )
            print("Vehículo no encontrado")

            campo_modelo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "model"))
            )
            campo_modelo.send_keys(modelo_esperado)
            campo_modelo.send_keys(Keys.ARROW_DOWN)
            campo_modelo.send_keys(Keys.ENTER)
            print(f"Se ingresó el modelo del vehículo: {modelo_esperado}")

            campo_marca = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "brand"))
            )
            time.sleep(1)
            campo_marca.send_keys(marca)
            time.sleep(1)
            campo_marca.send_keys(Keys.ARROW_DOWN)
            campo_marca.send_keys(Keys.ENTER)
            print(f"Se ingresó la marca del vehículo: {marca}")

            campo_linea = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "line"))
            )
            campo_linea.send_keys(linea)
            campo_linea.send_keys(Keys.ENTER)
            print(f"Se ingresó la línea del vehículo: {linea}")

            boton_siguiente2 = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button-stepper-next"]'))
                )
            boton_siguiente2.click()

            wait = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "card"))
            )
            cartas = driver.find_elements(By.CLASS_NAME, "card")
            cantidad_cartas = len(cartas)

            if cantidad_cartas == 1:
                print("Hay una sola carta. ")
            elif cantidad_cartas > 1:
                print(f"Se encontraron {cantidad_cartas} cartas.")
                print(f"Buscando código Fasecolda correspondiente...")

            from validar_fasecolda import filtrar_excel
            codigo_fasecolda = filtrar_excel()

            carta_encontrada = False
            for carta in cartas:
                codigo_elemento = carta.find_element(By.CSS_SELECTOR, 'h6[data-testid="identification-Fasecolda"]')
                codigo_en_carta = codigo_elemento.text.split(":")[-1].strip()

                codigo_en_carta_normalizado = codigo_en_carta.lstrip('0')

                if codigo_en_carta_normalizado == str(codigo_fasecolda):
                    boton = carta.find_element(By.CSS_SELECTOR, 'button[data-testid="button-is-my-vehicle"]')
                    boton.click()
                    print("Se seleccionó la carta correcta.")
                    carta_encontrada = True
                    break
            if not carta_encontrada:
                print(f"No se encontró una carta con el código Fasecolda {codigo_fasecolda}.")
        except Exception as e: 
            print(f"Error durante el proceso: {e}")

            try:
                wait = WebDriverWait(driver, 15)
                modelo_pantalla_elemento = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'h6[data-testid="identification-Modelo"]'))
                )
                modelo_pantalla = modelo_pantalla_elemento.text.split(":")[-1].strip()
                print(f"Modelo en pantalla: {modelo_pantalla}")
                print(f"Modelo esperado: {modelo_esperado}")

                if modelo_pantalla == modelo_esperado:
                    print(f"El modelo coincide. Se procede a continuar con el proceso.")

                    boton_es_mi_vehiculo = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button-is-my-vehicle"]'))
                    )
                    boton_es_mi_vehiculo.click()
                else:
                    print(f"El modelo no coincide.")
            except Exception as e:
                print(f"Error durante la verificación del modelo: {e}")

        try: 
            campo_ciudad = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ubication"))
            )
            campo_ciudad.send_keys(ciudad_circulacion)
            time.sleep(3)
            campo_ciudad.send_keys(Keys.ARROW_DOWN)
            campo_ciudad.send_keys(Keys.ENTER)
            print(f"Ciudad seleccionada: {ciudad_circulacion}")
            wait = WebDriverWait(driver, 5)

        except Exception as e:
            print(f"Error al seleccionar la ciudad: {e}")

        try:
            if servicio == "PUBLICO":
                boton_tipo_placa = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "plate_type"))
                )
                boton_tipo_placa.send_keys(Keys.ARROW_UP)
                boton_tipo_placa.send_keys(Keys.ENTER)
                print(f"Se ha cambiado el tipo de placa a: {servicio}")

                boton_tipo_uso = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "use_type"))
                )
                boton_tipo_uso.send_keys(Keys.ARROW_DOWN)
                boton_tipo_uso.send_keys(Keys.ARROW_DOWN)
                boton_tipo_uso.send_keys(Keys.ENTER)
                print(f"Se ha cambiado el tipo de uso")

            else:
                print(f"Tipo de placa: {servicio}")

            wait = WebDriverWait(driver, 3)

            boton_siguiente3 = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button-stepper-next"]'))
            )
            boton_siguiente3.click()
        except Exception as e:
            print(f"Error al seleccionar el tipo de placa")

        try:
            campo_nombre = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "first_name"))
                )
            borrar_y_escribir(campo_nombre, nombres)
            print(f"Se ingresó el nombre correctamente: {nombres}")

            campo_apellidos = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "last_name"))
            )
            borrar_y_escribir(campo_apellidos, apellidos)
            print(f"Se ingresó el apellido correctamente: {apellidos}")

            campo_fecha_nacimiento = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "birth_date"))
            )
            campo_fecha_nacimiento.click()
            campo_fecha_nacimiento.clear()
            campo_fecha_nacimiento.send_keys(fecha_nacimiento)
            print(f"Se ingresó la fecha de nacimiento correctamente: {fecha_nacimiento}")

            campo_correo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "email"))
            )
            campo_correo.send_keys(correo)
            time.sleep(1)

            if genero == "MASCULINO":
                boton_m = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//label[normalize-space(text())='Masculino']"))
                )
                boton_m.click()
                print("Se seleccionó el género masculino")

            elif genero == "FEMENINO":
                boton_f = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//label[normalize-space(text())='Femenino']"))
                )
                boton_f.click()
                print("Se seleccionó el género femenino")
            else:
                print(f"Valor inesperado para género: {genero} - {e}")

            if oneroso != "NO":
                campo_oneroso = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "onerous_beneficiary"))
                )
                campo_oneroso.click()
                print(f"Posee beneficiario oneroso.")

                seleccionar_oneroso = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "onerous_data"))
                )
                seleccionar_oneroso.send_keys(oneroso)
                seleccionar_oneroso.send_keys(Keys.ARROW_DOWN)
                seleccionar_oneroso.send_keys(Keys.ENTER)
                print(f"Se seleccionó el beneficiario oneroso: {oneroso}")
            else: 
                print(f"No posee beneficiario oneroso.")

            wait = WebDriverWait(driver, 10)
            boton_resumen = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button-stepper-next"]'))
            )
            boton_resumen.click()
        except Exception as e:
            print(f"Error al ingresar datos: {e}")

        '''try: 
            boton_confirmar_cotizar = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".k-button-solid-base"))
            )
            time.sleep(1)
            boton_confirmar_cotizar.click()
            print(f"Botón cancelar presionado")
        except Exception as e:
            print(f"No se pudo confirmar la cotización")
            '''
    finally:
        time.sleep(6)
        driver.quit()

if __name__ == "__main__":
    agente_motor()