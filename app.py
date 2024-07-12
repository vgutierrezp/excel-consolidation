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
def generate_excel_with_dates(df, store_name):
    output = BytesIO()
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Ejec.1']

    new_df = df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    # Convertir fechas a datetime y manejar errores
    new_df['Ult. Prev.'] = pd.to_datetime(new_df['Ult. Prev.'], format='%d/%m/%Y', errors='coerce')
    new_df['Ejec.1'] = pd.to_datetime(new_df['Ejec.1'], format='%d/%m/%Y', errors='coerce')

    # Filtrar filas con fechas no válidas en 'Ult. Prev.'
    new_df = new_df.dropna(subset=['Ult. Prev.'])

    # Obtener la fecha más reciente del último preventivo realizado para cada servicio
    new_df['Unique_Service'] = new_df['Familia'] + new_df['Tipo de Equipo'] + new_df['Tipo de Servicio'] + new_df['Ejecutor'] + new_df['Frecuencia'].astype(str)
    latest_preventives = new_df.loc[new_df.groupby('Unique_Service')['Ult. Prev.'].idxmax()]

    # Calcular fechas programadas
    for index, row in latest_preventives.iterrows():
        freq = row['Frecuencia']
        current_date = row['Ult. Prev.']
        col_num = 1
        while current_date <= max_date:
            latest_preventives.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        worksheet_name = 'Fechas Planificadas'
        latest_preventives.to_excel(writer, index=False, sheet_name=worksheet_name, startrow=2)
        worksheet = writer.sheets[worksheet_name]
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
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

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros con listas ordenadas
    month_order = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOV
