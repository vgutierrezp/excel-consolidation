import streamlit as st
import pandas as pd
from io import BytesIO
import os
from datetime import datetime

# Cargar el archivo consolidado desde el repositorio
@st.cache_data
def load_data():
    file_path = 'consolidated_file.xlsx'
    try:
        data = pd.read_excel(file_path)
        return data, file_path
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return pd.DataFrame(), file_path

# Función para convertir el DataFrame a Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    data, file_path = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Obtener fecha de última modificación
    last_modified = os.path.getmtime(file_path)
    last_modified_date = datetime.fromtimestamp(last_modified).strftime('%d/%m/%Y %H:%M:%S')

    # Mostrar fecha de última actualización
    st.write(f"Última actualización del archivo: {last_modified_date}")

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros con listas ordenadas
    months = sorted([''] + list(data['Mes'].dropna().unique()))
    brands = sorted([''] + list(data['Marca'].dropna().unique()))
    stores = sorted([''] + list(data['T
