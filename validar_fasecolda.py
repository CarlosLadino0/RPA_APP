import pandas as pd
import json
import os
import re

carpeta = "guia_fasecolda"
ruta_json = "datos/datos_extraidos.json"

def filtrar_excel(carpeta=carpeta, ruta_json=ruta_json):
    carpeta = "guia_fasecolda" 
    ruta_json = "datos/datos_extraidos.json"

    with open(ruta_json, "r", encoding="utf-8") as archivo_json:
        datos_json = json.load(archivo_json)
    archivo_excel = next((f for f in os.listdir(carpeta) if f.endswith('.xlsx')), None)
    if not archivo_excel:
        raise FileNotFoundError(f"No se encontró ningún archivo .xlsx en la carpeta {carpeta}")

    ruta_excel = os.path.join(carpeta, archivo_excel)
    df = pd.read_excel(ruta_excel, sheet_name="Codigos")

    columnas = ['Marca', 'Clase', 'Referencia1', 'Referencia3']
    for col in columnas:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    tarjeta = datos_json['tarjeta']
    marca = tarjeta['marca']
    clase_vehiculo = tarjeta['clase_vehiculo']
    linea = tarjeta['linea']
    cilindraje = str(tarjeta['cilindraje']).replace(".", "\\.")

    df_filtrado = df[
        (df['Marca'].str.contains(marca, case=False, na=False)) &
        (df['Clase'].str.contains(clase_vehiculo, case=False, na=False)) &
        (df['Referencia1'].str.contains(linea, case=False, na=False))  & 
        (df['Referencia3'].apply(lambda x: re.search(cilindraje, x) is not None)) 
    ]

    if df_filtrado.empty:
        raise ValueError("No se encontraron datos que coincidan con los filtros proporcionados")

    if len(df_filtrado) == 1:
        codigo_fasecolda = df_filtrado['Codigo'].iloc[0]  
    else:
        print("Múltiples resultados encontrados:")
        codigo_fasecolda = df_filtrado['Codigo'].tolist()
    return codigo_fasecolda

try:
    codigo_fasecolda = filtrar_excel()
    print("Codigo fasecolda encontrado:", codigo_fasecolda)
except Exception as e:
    print("Error:", e)