import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime, timedelta

# Función para generar el Excel con fechas calculadas
def generate_excel_with_dates(df, store_name):
    output = BytesIO()
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']

    new_df = df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    for index, row in new_df.iterrows():
        freq = row['Frecuencia']
        current_date = pd.to_datetime(row['Ult. Prev.'], errors='coerce')
        if pd.isnull(current_date):
            continue
        col_num = 1
        while current_date <= max_date:
            new_df.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        new_df.to_excel(writer, index=False, sheet_name='Fechas Planificadas', startrow=2)
        worksheet = writer.sheets['Fechas Planificadas']
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
    processed_data = output.getvalue()
    return processed_data

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

    # Verificar si se seleccionó una tienda
    selected_store = st.sidebar.selectbox('Selecciona una tienda para el Plan Anual de Mantenimiento', options=[''] + data['Tienda'].dropna().unique().tolist())
    
    # Botón para generar el Excel con fechas calculadas
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if not selected_store:
            st.sidebar.error("Selecciona una tienda para generar el Plan Anual de Mantenimiento")
        else:
            planned_excel_data = generate_excel_with_dates(filtered_data[filtered_data['Tienda'] == selected_store], selected_store)
            st.sidebar.download_button(
                label='Descargar Programa Anual de Mantenimiento',
                data=planned_excel_data,
                file_name='programa_anual_mantenimiento.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data.to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

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
