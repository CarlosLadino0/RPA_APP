import easyocr
import fitz
import cv2
import os
import re
import json
import numpy as np
import subprocess
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

reader = easyocr.Reader(['es'], gpu=False)

UPLOAD_FOLDER = 'datos/'

def extraer_texto_desde_pdf_o_imagen(file_path):
    """Función que maneja tanto PDFs como imágenes (JPG, PNG)."""
    if file_path.lower().endswith('.pdf'):
        return extraer_texto_desde_pdf(file_path)
    if file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
        return extraer_texto_desde_imagen(file_path)
    return ""

def extraer_texto_desde_pdf(pdf_path):
    """Extrae texto de todas las imágenes en el PDF."""
    texto_paginas = []
    temp_files = []
    try:
        with fitz.open(pdf_path) as pdf:
            for pagina in pdf:
                for img_index, img in enumerate(pagina.get_images(full=True)):
                    xref = img[0]
                    base_image = pdf.extract_image(xref)
                    image_bytes = base_image["image"]

                    temp_image_path = os.path.join(UPLOAD_FOLDER, f"temp_img_{img_index}.png")
                    temp_files.append(temp_image_path)

                    with open(temp_image_path, "wb") as image_file:
                        image_file.write(image_bytes)

                    texto_imagen = reader.readtext(temp_image_path, detail=0, paragraph=True)
                    texto_paginas.append(' '.join(texto_imagen))

        return ' '.join(texto_paginas)
    finally:
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.remove(temp_file)

def preprocesar_imagen(image_path):
    """Preprocesa la imagen para mejorar la precisión del OCR."""
    imagen = cv2.imread(image_path, cv2.IMREAD_COLOR)
    if imagen is None:
        raise ValueError(f"Error: No se pudo cargar la imagen en la ruta {image_path}")

    gris = cv2.cvtColor(imagen, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gris = clahe.apply(gris)
    gris = cv2.medianBlur(gris, 3)
    umbral = cv2.adaptiveThreshold(gris, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, 8)
    bordes = cv2.Canny(umbral, 50, 150)
    contornos, _ = cv2.findContours(bordes, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if contornos:
        contorno_max = max(contornos, key=cv2.contourArea)
        rect = cv2.minAreaRect(contorno_max)
        box = cv2.boxPoints(rect)
        box = np.int64(box)
        ancho, alto = rect[1]
        angulo = rect[2] if ancho < alto else rect[2] + 90

    preprocesada_path = os.path.join(UPLOAD_FOLDER, 'preprocesada.png')
    cv2.imwrite(preprocesada_path, gris)
    return preprocesada_path

def extraer_texto_desde_imagen(image_path):
    """Extrae texto de una imagen usando easyocr."""
    texto_imagen = reader.readtext(image_path, detail=0, paragraph=True)
    return ' '.join(texto_imagen)

def extraer_nombre_apellido(texto):
    """Extrae el nombre y apellido del texto extraído de la cédula."""
    nombres = ""
    apellidos = ""
    num_cedula = ""

    match_num_cedula = re.search(r'N[ÚU]MERO\s+(\d{1,3}(?:\.\d{3}){2,3})', texto)
    if match_num_cedula:
        num_cedula = match_num_cedula.group(1)
        print("Número de cédula extraído:", num_cedula)
    else: 
        print("No se pudo extraer el número de cédula")

    match_apellidos = re.search(r'([A-Z\s]+)\s+APELLIDOS', texto)
    if match_apellidos:
        apellidos = match_apellidos.group(1).strip()

    match_nombres = re.search(r'APELLIDOS\s([A-Z\s]+)\s+NOMBRES', texto)
    if match_nombres:
        nombres = match_nombres.group(1).strip()

    apellidos = " ".join(apellidos.split())
    nombres = " ".join(nombres.split())
    return apellidos, nombres, num_cedula

def extraer_datos_trasera(texto):
    """Extrae la fecha de nacimiento y fecha/lugar de expedición del texto."""
    fecha_nacimiento = ""
    fecha_lugar_expedicion = ""

    match_fecha_nacimiento = re.search(r'FECHA DE NACIMIENTO\s+(\d{1,2}-[A-Z]{3}-\d{4})', texto)
    if match_fecha_nacimiento:
        fecha_nacimiento = match_fecha_nacimiento.group(1)

    match_fecha_lugar = re.search(r'(\d{1,2}-[A-Z]{3}-\d{4}\s+[A-Z\s]+)\s+FECHA Y LUGAR DE EXPEDICI[ÓO]N', texto)
    if match_fecha_lugar:
        fecha_lugar_expedicion = match_fecha_lugar.group(1).strip()
    return fecha_nacimiento, fecha_lugar_expedicion

def extraer_datos_tarjeta_dinamicos(texto):
    datos_tarjeta = {
        'placa': '',
        'clase_vehiculo': '',
        'modelo': '',
        'servicio': '',
        'marca': '',
        'linea': '',
        'cilindraje': '',
    }

    texto = texto.replace(":", " ")
    texto = texto.replace("-", " ")

    match_placa = re.search(r'\b([A-Z]{3}\d{3}|[A-Z]{3}\d{2}[A-Z])\b', texto)
    if match_placa:
        datos_tarjeta['placa'] = match_placa.group(1)

    match_clase_vehiculo = re.search(r'\b(AUTOM[ÓO]VIL|BUS|BUSETA|CAMI[ÓO]N|CAMIONETA|CAMPERO|MICROBUS|TRACTOCAMI[ÓO]N|MOTOCICLETA|MOTOCARRO|MOTOTRICICLO|CUATRIMOTO|VOLQUETA)\b', 
    texto, re.IGNORECASE)
    if match_clase_vehiculo:
        datos_tarjeta['clase_vehiculo'] = match_clase_vehiculo.group(1).strip()

    match_servicio = re.search(r'\b(PARTICULAR|P[ÚU]BLICO|DIPLOM[ÁA]TICO|OFICIAL)\b', texto, re.IGNORECASE)
    if match_servicio:
        datos_tarjeta['servicio'] = match_servicio.group(1)

    match_modelo = re.search(r'\b(19[0-9]{2}|20[0-9]{2})\b', texto)
    if match_modelo:
        datos_tarjeta['modelo'] = match_modelo.group(1)

    match_marca = re.search(r'\b([A-Z]{3}\d{3}|[A-Z]{3}\d{2}[A-Z])\b\s+([A-Z]+)\b', texto)
    if match_marca:
        datos_tarjeta['marca'] = match_marca.group(2).strip()

    match_cilindraje = re.search(r'\b\d{3}\b|\b\d\.\d{3}\b', texto)
    if match_cilindraje:
        datos_tarjeta['cilindraje'] = match_cilindraje.group()

    if 'marca' in datos_tarjeta and 'modelo' in datos_tarjeta:
        match_linea = re.search(rf'\b{datos_tarjeta["marca"]}\s+(.*?)\s+\b{datos_tarjeta["modelo"]}\b', texto)
        if match_linea:
            datos_tarjeta['linea'] = match_linea.group(1).strip()
    return datos_tarjeta

def procesar_imagenes_en_paralelo(image_paths):
    """Procesa las imágenes en paralelo."""
    with ThreadPoolExecutor() as executor:
        resultados = list(executor.map(preprocesar_imagen, image_paths))
    return resultados

def procesar_documentos(cedula_path, tarjeta_path):
    texto_cedula = extraer_texto_desde_pdf_o_imagen(cedula_path)
    texto_tarjeta = extraer_texto_desde_pdf_o_imagen(tarjeta_path)
    apellidos, nombres, num_cedula = extraer_nombre_apellido(texto_cedula)
    fecha_nacimiento, fecha_lugar_expedicion = extraer_datos_trasera(texto_cedula)
    datos_tarjeta = extraer_datos_tarjeta_dinamicos(texto_tarjeta)

    datos_documentos = {
        'cedula': {
            'apellidos': apellidos,
            'nombres': nombres,
            'num_cedula': num_cedula,
            'fecha_nacimiento': fecha_nacimiento,
            'fecha_lugar_expedicion': fecha_lugar_expedicion
        },
        'tarjeta': datos_tarjeta,
        'texto_tarjeta': texto_tarjeta,
        'texto_cedula': texto_cedula
    }
    return datos_documentos