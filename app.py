import streamlit as st
import pandas as pd
import os

def generate_excel_with_dates(df, store):
    # Filtrar por tienda
    df = df[df['Tienda'] == store]

    # Verificar que las columnas necesarias existan
    required_columns = ['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ult. Prev.', 'Frecuencia']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Faltan las siguientes columnas necesarias en los datos: {', '.join(missing_columns)}")
        return None

    # Eliminar duplicados manteniendo el más reciente en "Ult. Prev."
    df = df.loc[df.groupby(['Familia', 'Tipo de Equipo', 'Tipo de Servicio'])['Ult. Prev.'].idxmax()]

    # Ordenar por las columnas especificadas
    df = df.sort_values(by=['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ult. Prev.'])

    # Inicializar un nuevo DataFrame para el programa anual
    program_df = pd.DataFrame()

    # Calcular las fechas programadas
    for index, row in df.iterrows():
        current_date = pd.to_datetime(row['Ult. Prev.'], errors='coerce')
        if pd.isna(current_date):
            continue
        
        program_dates = []
        while current_date < pd.Timestamp('2024-12-31'):
            current_date += pd.DateOffset(months=row['Frecuencia'])
            program_dates.append(current_date.strftime('%d/%m/%Y'))
        
        program_row = row.to_dict()
        for i, date in enumerate(program_dates, start=1):
            program_row[f'Prog.{i}'] = date
        
        program_df = program_df.append(program_row, ignore_index=True)
    
    return program_df

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

    # Botón para generar el programa anual de mantenimiento
    st.sidebar.header('Programa Anual de Mantenimiento')
    selected_store = st.sidebar.selectbox('Selecciona una tienda para su Plan Anual de Mantenimiento', options=[''] + sorted(data['Tienda'].dropna().unique().tolist()))

    if st.sidebar.button('Generar Plan Anual de Mantenimiento') and selected_store:
        planned_excel_data = generate_excel_with_dates(filtered_data, selected_store)
        if planned_excel_data is not None:
            planned_excel_data.to_excel('programa_anual_mantenimiento.xlsx', index=False)
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
