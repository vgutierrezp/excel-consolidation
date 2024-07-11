import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import os

# Cargar el archivo consolidado desde el repositorio
def load_data():
    data_file = 'consolidated_file.xlsx'
    if not os.path.exists(data_file):
        st.error(f"El archivo {data_file} no existe. Por favor, verifica la ruta y el nombre del archivo.")
        return pd.DataFrame()
    try:
        data = pd.read_excel(data_file)
        return data
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return pd.DataFrame()

# Función para convertir el DataFrame a Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# Función para generar el Excel con las fechas calculadas
def generate_excel_with_dates(df):
    output = BytesIO()
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Ejec.1']

    new_df = df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    # Convertir fechas a datetime y manejar errores
    new_df['Ult. Prev.'] = pd.to_datetime(new_df['Ult. Prev.'], format='%d/%m/%Y', errors='coerce')
    new_df['Ejec.1'] = pd.to_datetime(new_df['Ejec.1'], format='%d/%m/%Y', errors='coerce')

    # Calcular fechas programadas
    for index, row in new_df.iterrows():
        freq = row['Frecuencia']
        current_date = row['Ult. Prev.']
        col_num = 1
        while current_date <= max_date:
            new_df.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        worksheet_name = 'Fechas Planificadas'
        new_df.to_excel(writer, index=False, sheet_name=worksheet_name, startrow=2)
        worksheet = writer.sheets[worksheet_name]
        worksheet.write('A1', 'PLAN ANUAL DE MANTENIMIENTO')
        bold = writer.book.add_format({'bold': True})
        worksheet.set_row(0, None, bold)
        worksheet.set_column('H:H', None, writer.book.add_format({'num_format': 'dd/mm/yyyy'}))
    processed_data = output.getvalue()
    return processed_data

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    data = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Ordenar datos por Familia, Ejecutor, y Ult. Prev.
    data['Ult. Prev.'] = pd.to_datetime(data['Ult. Prev.'], errors='coerce').fillna(pd.Timestamp.min)
    data = data.sort_values(by=['Familia', 'Ejecutor', 'Ult. Prev.'])

    # Columnas a mostrar
    columns_to_show = ['Mes', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 
                       'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data_to_show = data[columns_to_show]

    # Mostrar los datos filtrados con las columnas seleccionadas
    st.write(data_to_show)

    # Opción para descargar el archivo filtrado
    st.sidebar.header('Descargar Datos')
    if not data.empty:
        excel_data = to_excel(data_to_show)
        st.sidebar.download_button(
            label='Descargar Excel',
            data=excel_data,
            file_name='filtered_data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    # Botón para generar el Excel con fechas calculadas
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        planned_excel_data = generate_excel_with_dates(data_to_show)
        st.sidebar.download_button(
            label='Descargar Programa Anual',
            data=planned_excel_data,
            file_name='programa_anual_mantenimiento.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
