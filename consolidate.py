import os
import pandas as pd

def consolidate_excels(source_folder, output_file):
    all_data = pd.DataFrame()

    # Recorrer todos los archivos en el directorio de proveedores
    for file in os.listdir(source_folder):
        if file.endswith(".xlsx") and file != 'consolidated_file.xlsx':  # Ignorar el archivo consolidado anterior
            file_path = os.path.join(source_folder, file)
            xls = pd.ExcelFile(file_path)

            # Recorrer todas las hojas del archivo Excel
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if 'Marca' in df.columns:
                    df = df[df['Marca'].notna() & (df['Marca'] != '')]  # Filtro sobre la columna "Marca"

                    # Agregar las columnas de origen
                    df['Archivo'] = file
                    df['Hoja'] = sheet_name
                    df['Fila'] = df.index + 2  # Asumiendo que el DataFrame original empieza en la fila 2 del archivo Excel

                    all_data = pd.concat([all_data, df], ignore_index=True)

    # Guardar el archivo consolidado
    all_data.to_excel(output_file, index=False)

if __name__ == "__main__":
    # Asegurarse de que la ruta de la carpeta de proveedores es correcta
    source_folder = r"C:\Users\vgutierrez\OneDrive - Servicios Compartidos de Restaurantes SAC\Documentos\01 Plan Preventivo Anual NGR\Preventivo\2024 PAM\PROVEEDORES"
    output_file = r"C:\Users\vgutierrez\chatbot_project\consolidated_file.xlsx"
    consolidate_excels(source_folder, output_file)
    print("Consolidación completada.")
