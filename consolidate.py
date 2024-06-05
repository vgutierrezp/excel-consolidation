import pandas as pd
import os

def consolidate_excels():
    folder_path = r'C:\Users\vgutierrez\OneDrive - Servicios Compartidos de Restaurantes SAC\Documentos\01 Plan Preventivo Anual NGR\Preventivo\2024 PAM\PROVEEDORES'
    consolidated_data = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx') and file != 'consolidated_file.xlsx':
                file_path = os.path.join(root, file)
                excel_data = pd.ExcelFile(file_path)

                for sheet_name in excel_data.sheet_names:
                    df = pd.read_excel(excel_data, sheet_name=sheet_name)
                    
                    # Remover fechas incorrectas y reemplazar con NaN
                    date_columns = ['Ult. Prev.', 'Plan.1', 'Prog.1', 'Ejec.1']
                    for col in date_columns:
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                        df[col] = df[col].where(df[col] >= '2023-01-01', pd.NaT)
                    
                    df['Source File'] = file
                    df['Sheet Name'] = sheet_name
                    consolidated_data.append(df)

    # Consolidar todos los datos en un solo DataFrame
    consolidated_df = pd.concat(consolidated_data, ignore_index=True)
    consolidated_df.to_excel('consolidated_file.xlsx', index=False)

# Llamar a la función de consolidación
consolidate_excels()
