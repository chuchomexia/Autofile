import fitz  # PyMuPDF
import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import os
import shutil
from transformers import AutoModelForCausalLM, AutoTokenizer

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        document = fitz.open(pdf_path)
        for page_num in range(len(document)):
            page = document[page_num]
            text += page.get_text()
        document.close()
    except Exception as e:
        print("Error al abrir o procesar el archivo PDF:", e)
    return text

def preprocess_text(text):
    # Eliminar caracteres especiales, saltos de línea y espacios en blanco innecesarios
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)
    # Convertir el texto a minúsculas
    text = text.lower()
    # Eliminar stopwords
    stop_words = set(stopwords.words('spanish'))
    word_tokens = word_tokenize(text)
    filtered_text = [word for word in word_tokens if word not in stop_words]
    return ' '.join(filtered_text)

def convert_to_unix_format(path):
    return path.replace('\\', '/')

def main():
    base_folder = input("¡Hola! Por favor ingresa la ruta de la carpeta que contiene los PDFs: ")
    base_folder = convert_to_unix_format(base_folder)

    # Crear la carpeta de destino si no existe
    txt_folder_name = 'txt_files'
    txt_folder = os.path.join(os.path.dirname(base_folder), txt_folder_name)

    # Si la carpeta ya existe, eliminarla y crearla de nuevo
    if os.path.exists(txt_folder):
        shutil.rmtree(txt_folder)
        print(f"La carpeta '{txt_folder_name}' existente ha sido eliminada.")
    os.makedirs(txt_folder)
    print(f"La carpeta '{txt_folder_name}' ha sido creada en '{os.path.dirname(txt_folder)}'.")

    # Ruta a la carpeta donde has extraído los archivos del modelo Mistral
    model_path = 'C:/Users/chuch/models/Mistral-7B-v0.2-Instruct'

    # Cargar el modelo y el tokenizador Mistral-7B
    model = AutoModelForCausalLM.from_pretrained(model_path)
    tokenizer = AutoTokenizer.from_pretrained(model_path)

    # Iterar sobre todos los archivos PDF en la carpeta
    for filename in os.listdir(base_folder):
        if filename.endswith('.pdf'):
            # Abrir el archivo PDF
            pdf_path = os.path.join(base_folder, filename)
            pdf_text = extract_text_from_pdf(pdf_path)

            # Preprocesamiento
            preprocessed_text = preprocess_text(pdf_text)

            # Utilizar el modelo Mistral-7B para procesar el texto
            input_ids = tokenizer.encode(preprocessed_text, return_tensors='pt')
            output = model.generate(input_ids, max_length=100, num_return_sequences=1, do_sample=True, top_k=50, top_p=0.95, num_beams=4)
            generated_text = tokenizer.decode(output[0], skip_special_tokens=True)

            # Nombre del archivo de texto de salida
            txt_filename = os.path.splitext(filename)[0] + '.txt'
            txt_path = os.path.join(txt_folder, txt_filename)

            # Guardar el texto procesado en un archivo .txt
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(generated_text)

    print('Extracción y procesamiento de texto completados.')

if __name__ == "__main__":
    main()