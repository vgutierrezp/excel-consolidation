import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

# Cargar el archivo consolidado desde el repositorio
def load_data():
    url = 'https://raw.githubusercontent.com/vgutierrezp/excel-consolidation/main/consolidated_file.xlsx'
    try:
        data = pd.read_excel(url)
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

# Función para generar el Excel filtrado
def generate_excel(data, store_name):
    output = BytesIO()
    columns_to_include = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Ejec.1']

    required_columns = ['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Tienda', 'Ejec.1', 'Ult. Prev.']
    missing_columns = [col for col in required_columns if col not in data.columns]

    if missing_columns:
        st.error(f"Faltan las siguientes columnas en los datos: {', '.join(missing_columns)}")
        return

    # Crear columna de concatenación única para identificar servicios únicos
    data['Unique_Service'] = data['Familia'] + data['Tipo de Equipo'] + data['Tipo de Servicio']

    # Filtrar los datos por tienda
    filtered_df = data[data['Tienda'] == store_name].copy()

    # Convertir columnas de fecha a datetime
    for col in ['Ejec.1', 'Ult. Prev.']:
        filtered_df[col] = pd.to_datetime(filtered_df[col], format='%Y-%m
