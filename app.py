import streamlit as st
import pandas as pd
import os

def generate_excel_with_dates(df, store_name):
    try:
        # Filtrar los datos por la tienda seleccionada
        df = df[df['Tienda'] == store_name]

        # Seleccionar las columnas relevantes
        columns = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']
        df = df[columns]

        # Ordenar por Familia, Tipo de Equipo, Tipo de Servicio y Ult. Prev.
        df = df.sort_values(by=['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ult. Prev.'])

        # Eliminar duplicados, manteniendo la fila con la fecha de 'Ult. Prev.' más reciente
        df = df.loc[df.groupby(['Familia', 'Tipo de Equipo', 'Tipo de Servicio'])['Ult. Prev.'].idxmax()]

        # Calcular las fechas programadas
        max_date = pd.Timestamp('2024-12-31')
        df_dates = df.copy()
        for i, row in df_dates.iterrows():
            current_date = row['Ult. Prev.']
            freq_months = row['Frecuencia']
            prog_dates = []
            while current_date <= max_date:
                prog_dates.append(current_date)
                current_date += pd.DateOffset(months=freq_months)
            for j, date in enumerate(prog_dates):
                df_dates.at[i, f'Prog.{j+1}'] = date

        # Crear el archivo Excel
        excel_file = f'programa_anual_mantenimiento_{store_name}.xlsx'
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            df_dates.to_excel(writer, sheet_name='Plan Anual', index=False, startrow=2)

            # Formato del archivo
            workbook = writer.book
            worksheet = writer.sheets['Plan Anual']
            title_format = workbook.add_format({'bold': True, 'font_size': 14})
            worksheet.write('A1', f'Programa Anual de Mantenimiento - {store_name}', title_format)

            # Formato de las columnas de fecha
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            for col_num, value in enumerate(df_dates.columns):
                if 'Prog' in value:
                    worksheet.set_column(col_num, col_num, 15, date_format)

        return excel_file

    except Exception as e:
        st.error(f"Error al generar el programa anual de mantenimiento: {str(e)}")
        return None

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

    # Botón para generar el Programa Anual de Mantenimiento
    st.sidebar.header('Generar Programa Anual de Mantenimiento')
    selected_store = st.sidebar.selectbox('Selecciona una tienda para su Plan Anual de Mantenimiento', sorted(data['Tienda'].unique()))
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if selected_store:
            planned_excel_data = generate_excel_with_dates(filtered_data, selected_store)
            if planned_excel_data:
                st.sidebar.markdown(f'[Descargar Programa Anual de Mantenimiento]({planned_excel_data})')
        else:
            st.warning('Selecciona una tienda para generar el Plan Anual de Mantenimiento.')

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
