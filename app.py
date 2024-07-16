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

    # Ordenar los meses cronológicamente
    months = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    data['Mes'] = pd.Categorical(data['Mes'], categories=months, ordered=True)

    # Definir los valores iniciales de los filtros
    mes = st.sidebar.selectbox('Mes', [''] + sorted(data['Mes'].dropna().unique().tolist(), key=lambda x: months.index(x)), key='mes')
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

    # Seleccionar columnas a mostrar
    selected_columns = ["Mes", "Llave1", "LlavePPto", "Ceco", "Marca", "Tienda", "Familia", "Tipo de Equipo", "Tipo de Servicio", "Ejecutor", "Frecuencia", "N° Equipos", "Ult. Prev."]
    filtered_data = filtered_data[selected_columns]

    # Mostrar la tabla filtrada
    st.dataframe(filtered_data)

    # Función para generar el plan anual de mantenimiento
    def generate_maintenance_plan(df, store_name):
        df['Unique_Service'] = df['Familia'] + df['Tipo de Equipo'] + df['Tipo de Servicio'] + df['Ejecutor'] + df['Frecuencia'].astype(str)
        plan_df = df.loc[df.groupby('Unique_Service')['Ult. Prev.'].idxmax()]

        # Calcular las fechas programadas
        for i in range(1, 13):
            plan_df[f'Prog.{i}'] = pd.to_datetime(plan_df['Ult. Prev.']) + pd.DateOffset(months=i * plan_df['Frecuencia'])

        # Seleccionar columnas para el plan
        plan_df = plan_df[["Tienda", "Familia", "Tipo de Equipo", "Tipo de Servicio", "Ejecutor", "N° Equipos", "Ult. Prev."] + [f'Prog.{i}' for i in range(1, 13)]]

        # Crear un archivo Excel
        excel_file = f'Plan_Anual_de_Mantenimiento_{store_name}.xlsx'
        with pd.ExcelWriter(excel_file) as writer:
            plan_df.to_excel(writer, sheet_name=store_name, index=False, startrow=2)
            workbook = writer.book
            worksheet = writer.sheets[store_name]
            worksheet.write(0, 0, f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
            worksheet.set_row(0, None, workbook.add_format({'bold': True}))

        return excel_file

    # Botón para generar el plan anual de mantenimiento
    st.sidebar.header('Generar Plan')
    tienda = st.sidebar.text_input("Nombre de la Tienda")
    if st.sidebar.button('Programa Anual de Mantenimiento'):
        if tienda:
            planned_excel_data = generate_maintenance_plan(filtered_data, tienda)
            st.sidebar.markdown(f'[Descargar Plan Anual de Mantenimiento]({planned_excel_data})')
        else:
            st.sidebar.error("Por favor, ingrese el nombre de la tienda.")

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
