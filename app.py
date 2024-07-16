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

    # Lista de meses en orden cronológico
    months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

    # Filtrar y ordenar los meses válidos
    valid_months = [m for m in data['Mes'].dropna().unique() if m in months]

    # Definir los valores iniciales de los filtros
    mes = st.sidebar.selectbox('Mes', [''] + sorted(valid_months, key=lambda x: months.index(x)), key='mes')
    marca = st.sidebar.selectbox('Marca', [''] + sorted(data['Marca'].dropna().unique()), key='marca')
    tienda = st.sidebar.selectbox('Tienda', [''] + sorted(data['Tienda'].dropna().unique()), key='tienda')
    familia = st.sidebar.selectbox('Familia', [''] + sorted(data['Familia'].dropna().unique()), key='familia')

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

    # Mostrar la tabla filtrada con las columnas específicas
    columns_to_show = ["Mes", "Llave1", "LlavePPto", "Ceco", "Marca", "Tienda", "Familia", "Tipo de Equipo", "Tipo de Servicio", "Ejecutor", "Frecuencia", "N° Equipos", "Ult. Prev."]
    st.dataframe(filtered_data[columns_to_show])

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data.to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

    # Sección para generar el plan anual de mantenimiento
    st.sidebar.header('Generar Plan')
    tienda = st.sidebar.text_input("Nombre de la Tienda")

    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if not tienda:
            st.warning("Por favor, ingrese el nombre de la tienda.")
        else:
            planned_excel_data = generate_excel_with_dates(filtered_data, tienda)
            st.sidebar.markdown(planned_excel_data)

def generate_excel_with_dates(df, store_name):
    # Definir las columnas a usar
    df = df[["Tienda", "Familia", "Tipo de Equipo", "Tipo de Servicio", "Ejecutor", "Frecuencia", "N° Equipos", "Ult. Prev."]]
    df['Unique_Service'] = df['Familia'] + df['Tipo de Equipo'] + df['Tipo de Servicio'] + df['Ejecutor'] + df['Frecuencia'].astype(str)

    # Filtrar por el servicio más reciente
    df['Ult. Prev.'] = pd.to_datetime(df['Ult. Prev.'], errors='coerce')
    df = df.sort_values(by='Ult. Prev.', ascending=False).drop_duplicates('Unique_Service')

    # Inicializar un nuevo DataFrame
    plan_df = df.copy()

    # Calcular las fechas programadas
    for i in range(1, 25):  # Calculando hasta 24 meses adelante
        plan_df[f'Prog.{i}'] = plan_df['Ult. Prev.'] + pd.DateOffset(months=i*plan_df['Frecuencia'].astype(int))

    # Crear el archivo Excel
    output_file = f"Plan Anual de Mantenimiento {store_name}.xlsx"
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(store_name)
        writer.sheets[store_name] = worksheet

        # Escribir el título y el DataFrame
        worksheet.write('A1', f"PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}")
        plan_df.to_excel(writer, sheet_name=store_name, startrow=2, index=False)

    return f'[Descargar Plan Anual de Mantenimiento](Plan Anual de Mantenimiento {store_name}.xlsx)'

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
