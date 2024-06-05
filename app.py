import streamlit as st
import pandas as pd
from io import BytesIO

# Cargar el archivo consolidado
@st.cache_data
def load_data():
    data = pd.read_excel('consolidated_file.xlsx')
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

    # Formatear las columnas de fecha y mantener las celdas vacías
    date_columns = ['Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    for col in date_columns:
        data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d/%m/%y')
        data[col] = data[col].replace('NaT', '')

    # Ordenar los meses según el calendario
    month_order = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    data['Mes'] = pd.Categorical(data['Mes'], categories=month_order, ordered=True)

    # Filtrado por columnas específicas
    st.sidebar.header('Filtros')

    # Inicializar filtros
    selected_month = st.sidebar.selectbox('Mes', options=[''] + month_order)
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    selected_brand = st.sidebar.selectbox('Marca', options=[''] + list(filtered_data['Marca'].dropna().unique()))
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    selected_store = st.sidebar.selectbox('Tienda', options=[''] + list(filtered_data['Tienda'].dropna().unique()))
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    selected_family = st.sidebar.selectbox('Familia', options=[''] + list(filtered_data['Familia'].dropna().unique()))
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    # Ordenar por Mes y Familia
    filtered_data = filtered_data.sort_values(by=['Mes', 'Familia'], ascending=[True, True])

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
