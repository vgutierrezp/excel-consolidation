import streamlit as st
import pandas as pd
from io import BytesIO
import os

# Cargar el archivo consolidado desde el repositorio
def load_data():
    file_url = 'https://raw.githubusercontent.com/vgutierrezp/excel-consolidation/main/consolidated_file.xlsx'
    try:
        data = pd.read_excel(file_url)
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

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    data = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros con listas ordenadas
    month_order = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    months = sorted(data['Mes'].dropna().unique(), key=lambda x: month_order.index(x))
    brands = sorted(data['Marca'].dropna().unique())
    stores = sorted(data['Tienda'].dropna().unique())
    families = sorted(data['Familia'].dropna().unique())

    # Crear filtros dependientes
    selected_month = st.sidebar.selectbox('Mes', options=[''] + months, key='mes')
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    selected_brand = st.sidebar.selectbox('Marca', options=[''] + sorted(filtered_data['Marca'].dropna().unique()), key='marca')
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    selected_store = st.sidebar.selectbox('Tienda', options=[''] + sorted(filtered_data['Tienda'].dropna().unique()), key='tienda')
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    selected_family = st.sidebar.selectbox('Familia', options=[''] + sorted(filtered_data['Familia'].dropna().unique()), key='familia')
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    # Columnas a mostrar
    columns_to_show = ['Mes', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 
                       'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data = filtered_data[columns_to_show]

    # Formatear las columnas de fecha
    date_columns = ['Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    for col in date_columns:
        data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d/%m/%y').fillna('')

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

if __name__ == "__main__":
    main()
