import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

# Cargar el archivo consolidado desde el repositorio
def load_data():
    url = 'https://raw.githubusercontent.com/vgutierrezp/excel-consolidation/main/consolidated_file.xlsx'
    try:
        data = pd.read_excel(url)
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

# Función para generar el Excel con las fechas calculadas
def generate_excel_with_dates(df, store_name):
    output = BytesIO()
    columns_to_copy = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']

    # Filtrar datos por Tienda
    df = df[df['Tienda'] == store_name]

    # Concatenar columnas
    df['Unique_Service'] = df['Familia'] + df['Tipo de Equipo'] + df['Tipo de Servicio']

    # Eliminar duplicados y quedarse con filas con fecha más reciente en Ejec.1 de enero al mes en curso
    current_month = datetime.now().month
    df['Ejec.1'] = pd.to_datetime(df['Ejec.1'], errors='coerce')
    df = df[(df['Ejec.1'].dt.month >= 1) & (df['Ejec.1'].dt.month <= current_month)]
    df = df.loc[df.groupby('Unique_Service')['Ejec.1'].idxmax()]

    # Buscar filas con fechas en meses posteriores en Ult. Prev.
    df_remaining = df[(df['Ejec.1'].dt.month > current_month) & (df['Ejec.1'].isnull())]
    df_remaining['Ult. Prev.'] = pd.to_datetime(df_remaining['Ult. Prev.'], errors='coerce')
    df_remaining = df_remaining.loc[df_remaining.groupby('Unique_Service')['Ult. Prev.'].idxmax()]

    # Concatenar ambos DataFrames
    new_df = pd.concat([df, df_remaining])

    new_df = new_df[columns_to_copy].copy()
    max_date = datetime(2024, 12, 31)

    for index, row in new_df.iterrows():
        freq = row['Frecuencia']
        current_date = row['Ult. Prev.']
        col_num = 1
        while current_date <= max_date:
            new_df.loc[index, f'Prog.{col_num}'] = current_date.strftime('%d/%m/%Y')
            current_date += timedelta(days=freq)
            col_num += 1

    new_df = new_df.dropna(subset=['Ult. Prev.'])

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        worksheet_name = 'Fechas Planificadas'
        new_df.to_excel(writer, index=False, sheet_name=worksheet_name, startrow=2)
        worksheet = writer.sheets[worksheet_name]
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
        bold = writer.book.add_format({'bold': True})
        worksheet.set_row(0, None, bold)
        worksheet.set_column('H:H', None, writer.book.add_format({'num_format': 'dd/mm/yyyy'}))
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

    # Botón para generar el Excel con fechas calculadas
    if selected_store:
        if st.sidebar.button('Programa Anual de Mantenimiento'):
            if selected_month or selected_brand or selected_family:
                st.sidebar.warning("Por favor, deje solo el filtro de tienda lleno.")
            else:
                planned_excel_data = generate_excel_with_dates(filtered_data, selected_store)
                st.sidebar.download_button(
                    label='Descargar Programa Anual',
                    data=planned_excel_data,
                    file_name=f'Plan de Mantenimiento Anual {selected_store}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
    else:
        st.sidebar.warning("Por favor, seleccione una tienda.")

if __name__ == "__main__":
    main()
