import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO

# Cargar el archivo consolidado desde el repositorio
def load_data():
    file_path = 'consolidated_file.xlsx'
    try:
        data = pd.read_excel(file_path)
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
def generate_excel_with_dates(df, store_name):
    output = BytesIO()
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']

    new_df = df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    for index, row in new_df.iterrows():
        freq = row['Frecuencia']
        current_date = pd.to_datetime(row['Ult. Prev.'], errors='coerce')
        col_num = 1
        while current_date <= max_date:
            new_df.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Escribir el título
        worksheet = writer.book.add_worksheet()
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
        new_df.to_excel(writer, index=False, sheet_name='Fechas Planificadas', startrow=2)
        writer.sheets['Fechas Planificadas'] = worksheet

        # Formato de celda para el título
        format_title = writer.book.add_format({'bold': True})
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}', format_title)

    processed_data = output.getvalue()
    return processed_data

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    data, file_path = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros con listas ordenadas
    month_order = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    months = sorted(data['Mes'].dropna().unique(), key=lambda x: (month_order.index(x) if x in month_order else float('inf')))
    brands = sorted([''] + list(data['Marca'].dropna().unique()))
    stores = sorted([''] + list(data['Tienda'].dropna().unique()))
    families = sorted([''] + list(data['Familia'].dropna().unique()))

    # Crear filtros dependientes
    selected_month = st.sidebar.selectbox('Mes', options=[''] + months)
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    selected_brand = st.sidebar.selectbox('Marca', options=[''] + sorted(filtered_data['Marca'].dropna().unique()))
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    selected_store = st.sidebar.selectbox('Tienda', options=[''] + sorted(filtered_data['Tienda'].dropna().unique()))
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    selected_family = st.sidebar.selectbox('Familia', options=[''] + sorted(filtered_data['Familia'].dropna().unique()))
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    # Columnas a mostrar
    columns_to_show = ['Mes', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 
                       'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data = filtered_data[columns_to_show]

    # Ordenar los meses según el calendario y luego por Familia
    data['Mes'] = pd.Categorical(data['Mes'], categories=month_order, ordered=True)
    data = data.sort_values(by=['Mes', 'Familia'], ascending=[True, True])

    # Mostrar los datos filtrados con las columnas seleccionadas
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
        if selected_store:
            planned_excel_data = generate_excel_with_dates(data, selected_store)
            st.sidebar.download_button(
                label='Descargar Programa Anual',
                data=planned_excel_data,
                file_name='programa_anual_mantenimiento.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.sidebar.error('Selecciona una tienda si desea su Plan Anual de Mantenimiento')

if __name__ == "__main__":
    main()
