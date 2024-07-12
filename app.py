import streamlit as st
import pandas as pd
import os

def main():
    st.title('Programa de Mantenimiento Preventivo')

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

    # Mostrar la tabla filtrada con columnas relevantes
    relevant_columns = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']
    st.dataframe(filtered_data[relevant_columns])

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data[relevant_columns].to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

    # Botón para generar el programa anual de mantenimiento
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if not tienda:
            st.warning('Por favor, selecciona una tienda para generar el Programa Anual de Mantenimiento.')
            return

        # Seleccionar las columnas relevantes
        program_data = filtered_data[relevant_columns].drop_duplicates()

        # Ordenar por 'Ult. Prev.' (ascendente) y 'Familia' (alfabético)
        program_data = program_data.sort_values(by=['Ult. Prev.', 'Familia'], ascending=[True, True])

        # Generar las fechas programadas
        max_date = pd.to_datetime('2024-12-31')
        for i, row in program_data.iterrows():
            last_date = pd.to_datetime(row['Ult. Prev.'], errors='coerce')
            if pd.isna(last_date):
                continue
            freq = row['Frecuencia']
            while last_date <= max_date:
                next_date = last_date + pd.DateOffset(months=freq)
                col_name = f'Prog.{(next_date.year - last_date.year) * 12 + (next_date.month - last_date.month)}'
                program_data.at[i, col_name] = next_date.strftime('%Y-%m-%d')
                last_date = next_date

        # Guardar el programa anual en un archivo Excel
        program_data.to_excel('programa_anual_mantenimiento.xlsx', index=False)

        # Proveer el link de descarga
        st.sidebar.markdown(f'[Descargar Programa Anual de Mantenimiento](programa_anual_mantenimiento.xlsx)')

# Ejecutar la aplicación
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
