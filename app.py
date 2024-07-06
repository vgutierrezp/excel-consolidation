import streamlit as st
import pandas as pd
from io import BytesIO
import os
from datetime import datetime, timedelta

# Cargar el archivo consolidado desde el repositorio
def load_data():
    file_path = 'consolidated_file.xlsx'
    try:
        data = pd.read_excel(file_path)
        st.write(f"Datos cargados correctamente desde {file_path}")
        return data, file_path
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return pd.DataFrame(), None

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
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']

    new_df = df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    for index, row in new_df.iterrows():
        freq = row['Frecuencia']
        current_date = pd.to_datetime(row['Ult. Prev.'], format='%d/%m/%y')
        col_num = 1
        while current_date <= max_date:
            new_df.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        new_df.to_excel(writer, index=False, sheet_name='Fechas Planificadas')
    processed_data = output.getvalue()
    return processed_data

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    data, file_path = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    st.write("Datos cargados:")
    st.write(data.head())

    # Verificar que las columnas existan
    required_columns = ['Mes', 'Marca', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    for col in required_columns:
        if col not in data.columns:
            st.error(f"La columna {col} no existe en los datos.")
            return

    st.write("Todas las columnas requeridas existen en los datos.")

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros con listas ordenadas
    month_order = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    months = sorted(data['Mes'].dropna().unique(), key=lambda x: (month_order.index(x) if x in month_order else float('inf')))
    st.write(f"Meses: {months}")
    brands = sorted([''] + list(data['Marca'].dropna().unique()))
    st.write(f"Marcas: {brands}")
    stores = sorted([''] + list(data['Tienda'].dropna().unique()))
    st.write(f"Tiendas: {stores}")
    families = sorted([''] + list(data['Familia'].dropna().unique()))
    st.write(f"Familias: {families}")

    # Crear filtros dependientes
    selected_month = st.sidebar.selectbox('Mes', options=[''] + months)
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    st.write(f"Datos filtrados por mes ({selected_month}):")
    st.write(filtered_data.head())

    selected_brand = st.sidebar.selectbox('Marca', options=[''] + sorted(filtered_data['Marca'].dropna().unique()))
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    st.write(f"Datos filtrados por marca ({selected_brand}):")
    st.write(filtered_data.head())

    selected_store = st.sidebar.selectbox('Tienda', options=[''] + sorted(filtered_data['Tienda'].dropna().unique()))
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    st.write(f"Datos filtrados por tienda ({selected_store}):")
    st.write(filtered_data.head())

    selected_family = st.sidebar.selectbox('Familia', options=[''] + sorted(filtered_data['Familia'].dropna().unique()))
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    st.write(f"Datos filtrados por familia ({selected_family}):")
    st.write(filtered_data.head())

    # Columnas a mostrar
    columns_to_show = ['Mes', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 
                       'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data = filtered_data[columns_to_show]

    st.write("Datos después de seleccionar columnas:")
    st.write(data.head())

    # Verificar si hay datos en las columnas de fecha antes de formatear
    for col in ['Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']:
        if col in data.columns:
            st.write(f"Formateando la columna: {col}")
            data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d/%m/%y').fillna('')
        else:
            st.write(f"La columna {col} no existe en los datos seleccionados.")

    # Ordenar los meses según el calendario y luego por Familia
    data['Mes'] = pd.Categorical(data['Mes'], categories=month_order, ordered=True)
    data = data.sort_values(by=['Mes', 'Familia'], ascending=[True, True])

    # Mostrar los datos filtrados con las columnas seleccionadas
    st.write("Datos finales para mostrar:")
    st.write(data)

    # Opción para descargar el archivo filtrado
    st.sidebar.header('Descargar Datos')
    if not data.empty:
        excel_data = to_excel(data)
        st.sidebar.download_button(
            label='Descargar Excel',
            data=excel_data,
            file_name='filtered_data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    # Botón para generar el Excel con fechas calculadas
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if selected_store == '':
            st.sidebar.error("Por favor, seleccione una tienda para generar el Programa Anual de Mantenimiento.")
        else:
            planned_excel_data = generate_excel_with_dates(filtered_data)
            st.sidebar.download_button(
                label='Descargar Programa Anual',
                data=planned_excel_data,
                file_name=f'programa_anual_mantenimiento_{selected_store}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()
