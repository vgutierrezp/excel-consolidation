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

    # Mostrar la tabla filtrada
    st.dataframe(filtered_data)

    # Función para generar el archivo Excel con el plan anual de mantenimiento
    def generate_annual_maintenance_plan(filtered_data):
        # Eliminar duplicados manteniendo la fila con la fecha más reciente en 'Ult. Prev.'
        filtered_data['Ult. Prev.'] = pd.to_datetime(filtered_data['Ult. Prev.'], errors='coerce')
        filtered_data = filtered_data.sort_values(by='Ult. Prev.').drop_duplicates(subset=['Familia', 'Tipo de Equipo', 'Tipo de Servicio'], keep='last')

        # Ordenar por 'Familia', 'Tipo de Equipo', 'Tipo de Servicio' y 'Ult. Prev.'
        filtered_data = filtered_data.sort_values(by=['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ult. Prev.'], ascending=[True, True, True, True])

        # Crear un nuevo DataFrame para el plan anual
        plan_anual = filtered_data.copy()

        # Calcular fechas programadas
        # Aquí puedes implementar el cálculo de las fechas programadas si es necesario

        # Guardar el plan anual en un archivo Excel
        plan_anual_file = 'programa_anual_mantenimiento.xlsx'
        with pd.ExcelWriter(plan_anual_file, engine='xlsxwriter') as writer:
            plan_anual.to_excel(writer, index=False, sheet_name='Plan Anual')
            workbook = writer.book
            worksheet = writer.sheets['Plan Anual']
            worksheet.write('A1', 'Programa Anual de Mantenimiento')
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 20)
            worksheet.set_column('D:D', 20)
            worksheet.set_column('E:E', 20)

        return plan_anual_file

    # Botón para generar el programa anual de mantenimiento
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        st.write("Generando el Programa Anual de Mantenimiento...")
        plan_anual_file = generate_annual_maintenance_plan(filtered_data)
        st.write(f"[Descargar Programa Anual de Mantenimiento]({plan_anual_file})")

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data.to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

# Ejecutar la aplicación
if 'mes' not in st.session_state:
    st.session_state['mes'] = ''
if 'marca' not in st.session_state:
    st.session_state['marca'] = ''
if 'tienda' not in st.session_state:
    st.session_state['tienda'] = ''
if 'familia' not in st.session_state:
    st.session_state['familia'] = ''

main()
