import os
import re
from PyPDF2 import PdfReader
from utils import config_data, get_quality_code

# Función para procesar archivos PDF
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
                    'Origen': config_data['origen'],
                    'Acceso': config_data['acceso'],
                    'Idioma': config_data['idioma'],
                    'Autor': config_data['autor'],
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
