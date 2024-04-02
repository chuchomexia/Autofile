import os
import re
import csv
import json
import subprocess
import datetime
import shutil
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# Función para convertir rutas a formato Unix
def convert_to_unix_format(path):
    return path.replace('\\', '/')

# Función para validar que la carpeta tiene el nombre correcto
def is_valid_base_folder(base_folder):
    required_folder_names = [
        "3140_C09.06_CONTRATO_PRESTACION_SERVICIOS",
        "3140_C09.08_ORDEN_COMPRA",
        "3140_C09.11_ORDEN_PRESTACION_SERVICIOS",
        "3140_C09.12_ORDEN_TRABAJO",
        "3140_C09.17_ORDEN_CONSULTORIA",
        "3140_C09.33_ORDEN_SUMINISTROS"
    ]
    
    base_folder_name = os.path.basename(base_folder)
    return base_folder_name in required_folder_names

# Función para obtener el código de calidad del archivo
def get_quality_code(file_name):
    pattern = r"F[A-Z]{2}.(\d{2,})"
    match = re.search(pattern, file_name)
    return match.group() if match else ""

# Función para cargar información estática desde un archivo CSV
def load_static_info():
    static_info = []
    with open('lists.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            static_info.append(row)
    return static_info

# Función para cargar equivalencias de nombres desde un archivo JSON
def load_name_equivalences():
    with open('equivalences_names.json', 'r') as f:
        return json.load(f)

# Función para cargar equivalencias de tipologías desde un archivo JSON
def load_typology_equivalences():
    with open('equivalences_typologies.json', 'r') as f:
        return json.load(f)

# Función para cargar equivalencias de subseries desde un archivo JSON
def load_subseries_equivalences():
    with open('equivalences_subseries.json', 'r') as f:
        return json.load(f)

# Función para obtener el código de unidad
def get_unit_code(base_folder_name):
    pattern = r"^\d{4}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "3140"

# Función para obtener el código de serie
def get_series_code(base_folder_name):
    pattern = r"C\d{2}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "C09"

# Función para obtener el nombre de expediente
def get_expedient_name(user_folder_name):
    return user_folder_name

# Función para obtener el código de subserie
def get_subseries_code(base_folder_name):
    pattern = r"C\d{2}.\d{2}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "C09.11"

# Función para obtener el nombre de la subserie
def get_subseries_name(base_folder_name):
    subseries_equivalences = load_subseries_equivalences()
    subseries_name = "Orden de Prestación de Servicios"

    pattern = r"C\d{2}.\d{2}"
    match = re.search(pattern, base_folder_name)
    if match:
        subseries_code = match.group()
        subseries_name = subseries_equivalences.get(subseries_code, "Orden de Prestación de Servicios")

    return subseries_name

# Función para obtener la descripción del contenido
def get_content_description(number):
    description = ""
    if number == 1:
        description = input("Documento con CC o NIT (CC XXXXXXXX): ")
    elif number == 2:
        description = "Nombre " + input("Nombre completo contratista: ")
    elif number == 3:
        description = "$ " + input("Valor total del contrato: ")
    return description

# Función para obtener la fecha del contrato
def get_contract_date(message):
    valid_date = False
    while not valid_date:
        date = input(message)
        try:
            date_obj = datetime.datetime.strptime(date, '%d%m%Y')
            valid_date = True
        except ValueError:
            print("¡Ups! Fecha ingresada no válida. Ingresa la fecha en formato DDMMAAAA.")
    return date_obj.strftime('%Y%m%d')

# Función para abrir un archivo PDF
def open_pdf(pdf_name):
    subprocess.Popen(['start', '', pdf_name], shell=True)

# Función para cerrar Adobe Acrobat
def close_pdf():
    with open(os.devnull, 'w') as devnull:
        subprocess.Popen(['taskkill', '/F', '/IM', 'Acrobat.exe'], stdout=devnull, stderr=devnull)

# Función para copiar el archivo Excel a la carpeta del usuario
def copy_excel_to_user_folder(user, base_folder, result_folder):
    user_folder = os.path.join(base_folder, user)
    if not os.path.exists(user_folder):
        os.makedirs(user_folder)
    
    excel_file_path = os.path.join(result_folder, f'{user}.xlsx')
    destination_path = os.path.join(user_folder, f'{user}.xlsx')
    
    shutil.copy(excel_file_path, destination_path)
    print(f"Archivo Excel guardado también en la carpeta del expediente {user}")
