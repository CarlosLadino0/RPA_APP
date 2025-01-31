import requests
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor

UPLOAD_FOLDER = 'datos/'
OCR_API_URL = 'https://api.ocr.space/parse/image'
API_KEY = '8ece7dc90288957'

def extraer_texto(file_path, lenguage='spa'):
    with open(file_path, 'rb') as file:
        response = requests.post(
            OCR_API_URL,
            files={'file': file},
            data={'apikey': API_KEY,
                  'language': lenguage, 
                  'isOverlayRequired': False,
                  'OCREngine': 2
                  }
        )
    if response.status_code == 200:
        resultado = response.json()
        if resultado.get("IsErroredOnProcessing"):
            raise ValueError(f"Error al procesar la imagen: {resultado.get('ErrorMessage')}")
        return ' '.join([r['ParsedText'] for r in resultado.get('ParsedResults', [])])
    else:
        raise ConnectionError(f"Error en la solicitud: {response.status_code}")

def extraer_texto_desde_pdf_o_imagen(file_path):
    if file_path.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png')):
        return extraer_texto(file_path)
    return ""

def guardar_texto_temporal(texto, nombre_archivo):
    """Guarda el texto en un archivo para analizar su estructura."""
    ruta_archivo = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    with open(ruta_archivo, 'w', encoding='utf-8') as archivo:
        archivo.write(texto)
    print(f"Texto guardado en {ruta_archivo}")

def extraer_datos_cedula(texto):
    numero_cedula_patron = r"(?:N[UÚ]MERO\s*([\d\.]{6,10})|([\d\.]{6,10})\s*N[UÚ]MERO|(?:^|\n)([\d\.]{6,10})\s*(?=\nN[UÚ]MERO))"
    apellidos_patron = r"([A-ZÁÉÍÓÚÑ ]+)(?=\sAPELLIDOS)"
    nombres_patron = r"([A-ZÁÉÍÓÚÑ ]+)(?=\sNOMBRES)"
    fecha_nacimiento_patron = r"(?:FECHA DE NACIMIENTO\s*|\n)(\d{2}-[A-Z]{3}-\d{4})"
    fecha_expedicion_patron = r"(\d{2}-[A-Z]{3}-\s?\d{4})\s*([A-ZÁÉÍÓÚÑ .]*)\s*(?=\s*FECHA Y LUGAR DE EXPEDICI[OÓ]N)"

    def normalizar_cedula(cedula):
        return cedula.replace('.', '') if cedula else None
    numero_cedula = re.search(numero_cedula_patron, texto, re.MULTILINE | re.IGNORECASE)
    if numero_cedula:
        cedula = next((grupo for grupo in numero_cedula.groups() if grupo), None)
    else:
        cedula = None
    apellidos = re.search(apellidos_patron, texto, re.IGNORECASE)
    nombres = re.search(nombres_patron, texto, re.IGNORECASE)
    fecha_nacimiento = re.search(fecha_nacimiento_patron, texto, re.IGNORECASE)
    fecha_expedicion = re.search(fecha_expedicion_patron, texto, re.IGNORECASE)

    datos = {
        "num_cedula": normalizar_cedula(cedula),
        "apellidos": apellidos.group(0).strip().upper() if apellidos else None,
        "nombres": nombres.group(0).strip().upper() if nombres else None,
        "fecha_nacimiento": fecha_nacimiento.group(1).strip() if fecha_nacimiento else None,
        "fecha_lugar_expedicion": None
    }

    if fecha_expedicion:
        fecha = fecha_expedicion.group(1).strip()
        ciudad = fecha_expedicion.group(2).strip().upper() if fecha_expedicion.group(2) else None
        datos["fecha_lugar_expedicion"] = f"{fecha} {ciudad}" if ciudad else fecha

    '''for campo, valor in datos.items():
        if valor is None:
            raise ValueError(f"Error: El campo '{campo}' no se extrajo correctamente")'''
    return datos

def extraer_datos_tp(texto):
    placa_patron = r"\b([A-Z]{3}\d{3}|[A-Z]{3}\d{2}[A-Z])\b"
    clase_vehiculo_patron = r"\b(AUTOM[ÓO]VIL|BUS|BUSETA|CAMI[ÓO]N|CAMIONETA|CAMPERO|MICROBUS|TRACTOCAMI[ÓO]N|MOTOCICLETA|MOTOCARRO|MOTOTRICICLO|CUATRIMOTO|VOLQUETA)\b"
    modelo_patron = r"\b(19[0-9]{2}|20[0-9]{2})\b"
    servicio_patron = r"\b(PARTICULAR|P[ÚU]BLICO|DIPLOM[ÁA]TICO|OFICIAL)\b"
    marca_patron = r"MARCA\s*([A-ZÁÉÍÓÚÑ]+)"
    linea_patron = r"L[IÍ]NEA\s*([A-ZÁÉÍÓÚÑ0-9\- ]+)"
    cilindraje_patron = r"\b\d{3}\b|\b\d\.\d{3}\b"

    placa = re.search(placa_patron, texto, re.IGNORECASE)
    clase_vehiculo = re.search(clase_vehiculo_patron, texto, re.IGNORECASE)
    modelo = re.search(modelo_patron, texto, re.IGNORECASE)
    servicio = re.search(servicio_patron, texto, re.IGNORECASE)
    marca = re.search(marca_patron, texto, re.IGNORECASE)
    linea = re.search(linea_patron, texto, re.IGNORECASE)  
    cilindraje = re.search(cilindraje_patron, texto, re.IGNORECASE)

    datos = {
        "placa": placa.group(1) if placa else None,
        "clase_vehiculo": clase_vehiculo.group(1).strip() if clase_vehiculo else None,
        "modelo": modelo.group(1) if modelo else None,
        "servicio": servicio.group(1) if servicio else None,
        "marca": marca.group(1).strip().upper() if marca else None,
        "linea": linea.group(1).strip().upper() if linea else None, 
        "cilindraje": cilindraje.group() if cilindraje else None,
    }
    return datos

def procesar_documentos(cedula_path, tarjeta_path):
    texto_cedula = extraer_texto_desde_pdf_o_imagen(cedula_path)
    texto_tarjeta = extraer_texto_desde_pdf_o_imagen(tarjeta_path)

    guardar_texto_temporal(texto_cedula, "texto_cedula.txt")
    guardar_texto_temporal(texto_tarjeta, "texto_tarjeta.txt")

    datos_tarjeta = extraer_datos_tp(texto_tarjeta)
    datos_cedula = extraer_datos_cedula(texto_cedula)

    datos_documentos = {
        'texto_cedula': texto_cedula,
        'texto_tarjeta': texto_tarjeta,
        'datos_cedula': datos_cedula,
        'datos_tarjeta': datos_tarjeta
    }
    return datos_documentos