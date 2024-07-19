import pandas as pd
import os

def consolidate_excels():
    # Directorio de entrada y salida
    input_directory = 'C:/Users/vgutierrez/chatbot_project/excel_files'
    output_file = 'C:/Users/vgutierrez/chatbot_project/consolidated_file.xlsx'

    # Lista para almacenar los DataFrames
    consolidated_data = []

    # Recorrer todos los archivos Excel en el directorio de entrada
    for filename in os.listdir(input_directory):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(input_directory, filename)
            # Leer el archivo Excel
            df = pd.read_excel(file_path)
            consolidated_data.append(df)

    # Concatenar todos los DataFrames
    consolidated_df = pd.concat(consolidated_data, ignore_index=True)

    # Convertir 'Ult. Prev.' a formato de fecha y 'Frecuencia' a numérico
    consolidated_df['Ult. Prev.'] = pd.to_datetime(consolidated_df['Ult. Prev.'], errors='coerce')
    consolidated_df['Frecuencia'] = pd.to_numeric(consolidated_df['Frecuencia'], errors='coerce')

    # Eliminar filas duplicadas
    consolidated_df = consolidated_df.drop_duplicates()

    # Guardar el DataFrame consolidado en un archivo Excel
    consolidated_df.to_excel(output_file, index=False)
    print("Consolidation complete.")

if __name__ == "__main__":
    consolidate_excels()
