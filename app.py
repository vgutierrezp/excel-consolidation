import streamlit as st
import pandas as pd
import os

def main():
    st.title('Consolidación de Archivos Excel')

    # Ruta del archivo Excel
    data_file = 'consolidated_file.xlsx'

    # Verificar si el archivo existe
    if not os.path.exists(data_file):
        st.error(f"El archivo {data_file} no existe. Por favor, verifica la ruta y el nombre del archivo.")
        return

    try:
        # Cargar los datos desde el archivo Excel
        data = pd.read_excel(data_file)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {str(e)}")
        return

    # Definir los filtros
    st.sidebar.header('Filtros')

    # Botón para limpiar los filtros
    if st.sidebar.button('Limpiar Filtros'):
        st.session_state['mes'] = ''
        st.session_state['marca'] = ''
        st.session_state['tienda'] = ''
        st.session_state['familia'] = ''

    # Definir los valores iniciales de los filtros
    mes = st.sidebar.selectbox('Mes', [''] + sorted(data['Mes'].dropna().unique().tolist()), key='mes')
    marca = st.sidebar.selectbox('Marca', [''] + sorted(data['Marca'].dropna().unique().tolist()), key='marca')
    tienda = st.sidebar.selectbox('Tienda', [''] + sorted(data['Tienda'].dropna().unique().tolist()), key='tienda')
    familia = st.sidebar.selectbox('Familia', [''] + sorted(data['Familia'].dropna().unique().tolist()), key='familia')

    # Aplicar los filtros a los datos
    filtered_data = data.copy()
    if mes:
        filtered_data = filtered_data[filtered_data['Mes'] == mes]
    if marca:
        filtered_data = filtered_data[filtered_data['Marca'] == marca]
    if tienda:
        filtered_data = filtered_data[filtered_data['Tienda'] == tienda]
    if familia:
        filtered_data = filtered_data[filtered_data['Familia'] == familia]

    # Convertir fechas a datetime y manejar errores
    filtered_data['Ult. Prev.'] = pd.to_datetime(filtered_data['Ult. Prev.'], errors='coerce').fillna(pd.Timestamp.min)

    # Ordenar datos por Familia, Ejecutor, y Ult. Prev.
    filtered_data = filtered_data.sort_values(by=['Familia', 'Ejecutor', 'Ult. Prev.'])

    # Columnas a mostrar
    columns_to_show = ['Mes', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 
                       'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data_to_show = filtered_data[columns_to_show]

    # Mostrar los datos filtrados con las columnas seleccionadas
    st.write(data_to_show)

if __name__ == "__main__":
    if 'mes' not in st.session_state:
        st.session_state['mes'] = ''
    if 'marca' not in st.session_state:
        st.session_state['marca'] = ''
    if 'tienda' not in st.session_state:
        st.session_state['tienda'] = ''
    if 'familia' not in st.session_state:
        st.session_state['familia'] = ''
    main()
