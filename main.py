import os
import functions

def main():
    base_folder = input("¡Hola! Por favor ingresa la carpeta de contratos con los que trabajaremos hoy: ")
    base_folder = functions.convert_to_unix_format(base_folder)

    if not functions.check_base_folder(base_folder):
        print("¡Ups! La carpeta que ingresaste parece ser la incorrecta :c Deberías revisarla y estar más pendiente de tu trabajo.")
        return

    result_folder = os.path.join(base_folder, 'Excel')
    if not os.path.exists(result_folder):
        os.makedirs(result_folder)

    name_equivalences = functions.load_name_equivalences()
    typology_equivalences = functions.load_typology_equivalences()
    static_info = functions.load_static_info()

    initial_order_input = input("Orden del expediente: ")
    if initial_order_input.strip():
        initial_order = int(initial_order_input)
    else:
        initial_order = 1

    for user in sorted(os.listdir(base_folder)):
        user_folder = os.path.join(base_folder, user)
        if not os.path.isdir(user_folder) or user == 'Excel':
            continue

        if os.path.isdir(user_folder) and user != 'Excel':
            pdf_name = None
            for pdf_file in os.listdir(user_folder):
                if "FCO.66_Acta_" in pdf_file and pdf_file.endswith(".pdf"):
                    pdf_name = os.path.join(user_folder, pdf_file)
                    break

            if pdf_name:
                functions.open_pdf(pdf_name)
            else:
                print("No se encontró ningún PDF para abrir.")

            analyze_expedient = input(f"¿Deseas analizar el expediente '{user}'? (S/N): ")
            if analyze_expedient.lower() != 's':
                print(f"El expediente '{user}' será omitido.")
                continue

            print(f"Iniciando diligenciamiento del expediente {initial_order} {user}...")

            df, df_expedient = functions.generate_user_data_frame(user, base_folder, name_equivalences, typology_equivalences, initial_order)

            if df is not None and df_expedient is not None:
                functions.save_to_excel(user, user_folder, result_folder, df, df_expedient)
                functions.copy_excel_to_user_folder(user, base_folder, result_folder)

            initial_order += 1

            expedient_marked = False

            while True:
                response = input("¿Lo revisamos luego? (S/N): ")
                if response.lower() == 's':
                    with open(os.path.join(base_folder, 'expedientes_por_revisar.txt'), 'a') as report:
                        report.write(user + '\\n')
                    expedient_marked = True
                    break
                elif response.lower() == 'n':
                    break
                else:
                    print("¡Ups! Respuesta no válida. Por favor ingresa 'S' para marcarlo o 'N' para omitir.")

            # response = input("¿Continuamos? (S/N): ")
            # if response.lower() != 's':
            #     break

if __name__ == "__main__":
    main()
