import streamlit as st
import pandas as pd
from io import BytesIO
import os

# Cargar el archivo consolidado
@st.cache_data
def load_data():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, 'consolidated_file.xlsx')
    data = pd.read_excel(file_path)
    return data

# Función para convertir el DataFrame a Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# Función principal
def main():
    st.title("PLAN ANUAL DE MANTENIMIENTO PREVENTIVO")

    data = load_data()

    # Verificar si hay duplicados y eliminarlos
    data = data.drop_duplicates()

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros
    selected_month = st.sidebar.selectbox('Mes', options=[''] + list(data['Mes'].dropna().unique()))
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    selected_brand = st.sidebar.selectbox('Marca', options=[''] + list(filtered_data['Marca'].dropna().unique()))
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    selected_store = st.sidebar.selectbox('Tienda', options=[''] + list(filtered_data['Tienda'].dropna().unique()))
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    selected_family = st.sidebar.selectbox('Familia', options=[''] + list(filtered_data['Familia'].dropna().unique()))
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    # Mostrar los datos filtrados
    st.write(filtered_data)

    # Opción para descargar el archivo filtrado
    st.sidebar.header('Descargar Datos')
    if not filtered_data.empty:
        excel_data = to_excel(filtered_data)
        st.sidebar.download_button(
            label='Descargar Excel',
            data=excel_data,
            file_name='filtered_data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
