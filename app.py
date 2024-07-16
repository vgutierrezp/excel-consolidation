import streamlit as st
import pandas as pd
import os
import io

def generate_excel_with_dates(df, store_name):
    # Concatenar las columnas relevantes para identificar servicios únicos
    df['Unique_Service'] = df['Familia'] + '_' + df['Tipo de Equipo'] + '_' + df['Tipo de Servicio'] + '_' + df['Ejecutor'] + '_' + df['Frecuencia'].astype(str)
    
    # Convertir 'Ult. Prev.' a formato de fecha y manejar errores
    df['Ult. Prev.'] = pd.to_datetime(df['Ult. Prev.'], errors='coerce')
    
    # Ordenar y filtrar duplicados
    df = df.sort_values(by=['Ult. Prev.'], ascending=False)
    df = df.loc[df.groupby('Unique_Service')['Ult. Prev.'].idxmax()]
    
    # Crear el DataFrame para el plan anual
    plan_df = df[['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'N° Equipos', 'Ult. Prev.']]
    
    # Calcular las fechas programadas
    for i in range(1, 13):  # Ajustar el rango según la cantidad de frecuencias necesarias
        plan_df[f'Prog.{i}'] = plan_df['Ult. Prev.'] + pd.DateOffset(months=i*plan_df['Frecuencia'])
    
    # Formatear las fechas
    date_columns = [col for col in plan_df.columns if 'Prog.' in col]
    for col in date_columns:
        plan_df[col] = plan_df[col].dt.strftime('%d/%m/%Y')
    
    # Crear el archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        plan_df.to_excel(writer, index=False, sheet_name=store_name)
        workbook = writer.book
        worksheet = writer.sheets[store_name]
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}', workbook.add_format({'bold': True}))
        for col_num, value in enumerate(plan_df.columns.values):
            worksheet.write(2, col_num, value, workbook.add_format({'bold': True}))
        writer.save()
    
    return output.getvalue()

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

    # Definir los valores iniciales de los filtros
    # Ordenar meses cronológicamente
    months = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    unique_months = [m for m in data['Mes'].dropna().unique() if m in months]
    mes = st.sidebar.selectbox('Mes', [''] + sorted(unique_months, key=lambda x: months.index(x)), key='mes')
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

    # Convertir 'Ult. Prev.' a formato de fecha y manejar errores
    filtered_data['Ult. Prev.'] = pd.to_datetime(filtered_data['Ult. Prev.'], errors='coerce')

    # Ordenar la tabla de manera cronológica
    filtered_data = filtered_data.sort_values(by=['Ult. Prev.'], ascending=True)

    # Mostrar la tabla filtrada
    st.dataframe(filtered_data)

    # Botón para descargar el archivo filtrado en Excel
    st.sidebar.header('Descargar Datos')
    if st.sidebar.button('Descargar Excel'):
        filtered_data.to_excel('filtered_data.xlsx', index=False)
        st.sidebar.markdown(f'[Descargar archivo filtrado](filtered_data.xlsx)')

    # Botón para generar el plan anual de mantenimiento
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if tienda:
            planned_excel_data = generate_excel_with_dates(filtered_data, tienda)
            st.sidebar.download_button(label='Descargar Plan Anual de Mantenimiento', data=planned_excel_data, file_name=f'Plan Anual de Mantenimiento {tienda}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.warning('Selecciona una tienda si desea su Plan Anual de Mantenimiento')

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
