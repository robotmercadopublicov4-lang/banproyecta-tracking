import os
import re
import zipfile
import requests
import pandas as pd
from datetime import datetime
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


def descargar_licitaciones(mes: int, anio: int):

    print(f"\nDescargando licitaciones {anio}-{mes}...")

    # URL oficial
    url = f"https://transparenciachc.blob.core.windows.net/lic-da/{anio}-{mes}.zip"

    # Carpetas
    raw_path = "data/raw"
    output_path = "data/output"

    os.makedirs(raw_path, exist_ok=True)
    os.makedirs(output_path, exist_ok=True)

    # Archivo ZIP
    zip_filename = f"{anio}-{mes}.zip"
    zip_path = os.path.join(raw_path, zip_filename)

    # Descargar ZIP
    response = requests.get(url)

    if response.status_code != 200:
        print("ERROR descargando archivo")
        return None

    with open(zip_path, "wb") as file:
        file.write(response.content)

    print("ZIP descargado correctamente")

    # Descomprimir
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(raw_path)

    print("ZIP descomprimido")

    # Buscar CSV
    csv_filename = f"lic_{anio}-{mes}.csv"
    csv_path = os.path.join(raw_path, csv_filename)

    if not os.path.exists(csv_path):
        print("ERROR: CSV no encontrado")
        return None

    print("Leyendo CSV...")

    # Leer CSV
    df = pd.read_csv(
        csv_path,
        sep=';',
        encoding='ISO-8859-1',
        low_memory=False
    )

    print(f"Total registros: {len(df)}")

    return df


def filtrar_licitaciones(
        df,
        monto_minimo=0,
        estado='Adjudicada'
):

    print("\nFiltrando licitaciones...")

    # Validar columnas
    columnas_necesarias = [
        'CodigoExterno',
        'NombreProveedor',
        'RutProveedor',
        'MontoLineaAdjudica',
        'Estado',
        'RegionUnidad',
        'CantidadReclamos'
    ]

    for col in columnas_necesarias:
        if col not in df.columns:
            print(f"Falta columna: {col}")
            return None

    # Limpiar monto
    df['MontoLineaAdjudica'] = (
        df['MontoLineaAdjudica']
        .astype(str)
        .str.replace(',', '.', regex=False)
    )

    df['MontoLineaAdjudica'] = pd.to_numeric(
        df['MontoLineaAdjudica'],
        errors='coerce'
    )

    # Filtro
    df_filtrado = df[
        (df['Estado'] == estado) &
        (df['MontoLineaAdjudica'] >= monto_minimo)
    ]

    print(f"Licitaciones filtradas: {len(df_filtrado)}")

    return df_filtrado


def limpiar_texto_excel(texto):

    if not isinstance(texto, str):
        return texto

    # Remueve caracteres ilegales Excel
    texto = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', texto)

    return texto


def exportar_excel(df, nombre_archivo):

    output_path = "data/output"
    os.makedirs(output_path, exist_ok=True)

    archivo_salida = os.path.join(output_path, nombre_archivo)

    # Limpiar dataframe completo
    for col in df.columns:
        df[col] = df[col].apply(limpiar_texto_excel)

    # Exportar
    df.to_excel(archivo_salida, index=False)

    print(f"\nExcel exportado:")
    print(archivo_salida)