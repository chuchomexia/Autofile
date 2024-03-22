import os
import re
import csv
import json
import shutil
import datetime
import subprocess
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# Funciones de utilidad

def convert_to_unix_format(path):
    return path.replace('\\', '/')

def load_name_equivalences():
    with open('equivalences_names.json', 'r') as f:
        return json.load(f)

def load_typology_equivalences():
    with open('equivalences_typologies.json', 'r') as f:
        return json.load(f)

def load_subseries_equivalences():
    with open('equivalences_subseries.json', 'r') as f:
        return json.load(f)

def load_static_info():
    static_info = []
    with open('lists.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            static_info.append(row)
    return static_info

def get_quality_code(file_name):
    pattern = r"F[A-Z]{2}.(\d{2,})"
    match = re.search(pattern, file_name)
    return match.group() if match else ""

def get_unit_code(base_folder_name):
    pattern = r"^\d{4}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "3140"

def get_series_code(base_folder_name):
    pattern = r"C\d{2}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "C09"

def get_subseries_code(base_folder_name):
    pattern = r"C\d{2}.\d{2}"
    match = re.search(pattern, base_folder_name)
    return match.group() if match else "C09.11"

def get_unit_name():
    return "División Financiera"

def get_series_name():
    return "Contratos"

def get_subseries_name(base_folder_name):
    subseries_equivalences = load_subseries_equivalences()
    subseries_name = "Orden de Prestación de Servicios"

    pattern = r"C\d{2}.\d{2}"
    match = re.search(pattern, base_folder_name)
    if match:
        subseries_code = match.group()
        subseries_name = subseries_equivalences.get(subseries_code, "Orden de Prestación de Servicios")

    return subseries_name

def get_content_description(number):
    description = ""
    if number == 1:
        description = input("Documento con CC o NIT (CC XXXXXXXX): ")
    elif number == 2:
        description = "Nombre " + input("Nombre completo contratista: ")
    elif number == 3:
        description = "$ " + input("Valor total del contrato: ")
    return description

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

def open_pdf(pdf_name):
    subprocess.Popen(['start', '', pdf_name], shell=True)

def close_pdf():
    with open(os.devnull, 'w') as devnull:
        subprocess.Popen(['taskkill', '/F', '/IM', 'Acrobat.exe'], stdout=devnull, stderr=devnull)

# Funciones para el procesamiento de archivos PDF

def process_pdf(user_folder, document_order, previous_main_file_date, name_equivalences, typology_equivalences, static_info):
    pdf_info = []
    start_page = 1
    previous_end_page = 0

    for pdf_file in sorted(os.listdir(user_folder)):
        if not pdf_file.endswith('.pdf'):
            continue

        full_path = os.path.join(user_folder, pdf_file)
        try:
            with open(full_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                pages = len(pdf_reader.pages)

                start_page = previous_end_page + 1
                end_page = start_page + pages - 1

                file_date = None
                date_match = re.search(r"_(\d{8})_", pdf_file)
                if date_match:
                    file_date = date_match.group(1)

                if "Anexo" in pdf_file:
                    if previous_main_file_date:
                        file_date = previous_main_file_date
                    else:
                        raise ValueError(f"No se han diligenciado correctamente las fechas en el expediente {pdf_file}")
                else:
                    previous_main_file_date = file_date

                document_order += 1

                document_name = name_equivalences.get(pdf_file, '')
                typology_document = typology_equivalences.get(pdf_file, '')
                pdf_info.append({
                    'Nombre del archivo': pdf_file,
                    'Nombre del documento': document_name,
                    'Tipología documental': typology_document,
                    'Fecha de creación del documento': file_date,
                    'Fecha incorporación expediente': file_date,
                    'Orden documento expediente': document_order,
                    'Página inicio': start_page,
                    'Página fin': end_page,
                    'Origen': origen,
                    'Acceso': acceso,
                    'Idioma': idioma,
                    'Autor': autor,
                    'Código calidad': get_quality_code(pdf_file),
                    'Numero': '',
                    'Año': '',
                    'Metadato 1': '',
                    'Metadato 2': ''
                })

                previous_end_page = end_page

        except Exception as e:
            print(f"No se pudo leer el archivo {pdf_file}: {str(e)}")

    return pdf_info, document_order, previous_main_file_date
pass

def update_value(row, name_equivalences, typology_equivalences):
    # La lógica de actualización de valores se mantiene aquí
    pass

# Otras funciones de procesamiento de datos y generación de marco de datos

def generate_user_data_frame(user, base_folder, name_equivalences, typology_equivalences, initial_order=1):
    # La lógica de generación de marco de datos se mantiene aquí
    pass

def save_to_excel(user, user_folder, result_folder, df, df_expedient):
    # La lógica de guardado en Excel se mantiene aquí
    pass

def copy_excel_to_user_folder(user, base_folder, result_folder):
    # La lógica de copia de archivo Excel se mantiene aquí
    pass
