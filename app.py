import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO

def generate_excel_with_dates(df, store=None):
    # Mantener solo la fila con la fecha más reciente en 'Ult. Prev.' para cada combinación de 'Familia', 'Tipo de Equipo' y 'Tipo de Servicio'
    df = df.sort_values(by=['Ult. Prev.'], ascending=False).drop_duplicates(subset=['Familia', 'Tipo de Equipo', 'Tipo de Servicio'], keep='first')
    
    # Filtrar por tienda si se proporciona
    if store:
        df = df[df['Tienda'] == store]

    # Generar fechas programadas
    program_cols = ['Prog.1', 'Prog.2', 'Prog.3', 'Prog.4', 'Prog.5', 'Prog.6', 'Prog.7', 'Prog.8', 'Prog.9', 'Prog.10', 'Prog.11', 'Prog.12']
    for i, row in df.iterrows():
        current_date = pd.to_datetime(row['Ult. Prev.'])
        frequency = row['Frecuencia']
        for j in range(12):
            current_date += pd.DateOffset(months=frequency)
            df.at[i, program_cols[j]] = current_date if current_date.year <= 2024 else None

    # Crear el archivo Excel
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Programa Anual')

    # Formatear el archivo Excel
    workbook = writer.book
    worksheet = writer.sheets['Programa Anual']
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    for col in range(len(df.columns) - len(program_cols), len(df.columns)):
        worksheet.set_column(col, col, 15, date_format)

    writer.save()
    output.seek(0)
    return output

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

    # Mostrar la tabla filtrada
    st.dataframe(filtered_data)

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data.to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

    st.sidebar.header('Programa Anual de Mantenimiento')
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        stores = filtered_data['Tienda'].unique().tolist()
        if len(stores) > 1:
            selected_store = st.selectbox('Selecciona una tienda si desea su Plan Anual de Mantenimiento', stores)
        else:
            selected_store = stores[0]
        planned_excel_data = generate_excel_with_dates(filtered_data, selected_store)
        st.sidebar.download_button(
            label='Descargar
