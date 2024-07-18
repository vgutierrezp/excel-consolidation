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

# Función para generar el Excel filtrado
def generate_excel(data, store_name):
    output = BytesIO()
    columns_to_include = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Ejec.1']

    required_columns = ['Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Tienda', 'Ejec.1', 'Ult. Prev.']
    missing_columns = [col for col in required_columns if col not in data.columns]

    if missing_columns:
        st.error(f"Faltan las siguientes columnas en los datos: {', '.join(missing_columns)}")
        return

    # Crear columna de concatenación única para identificar servicios únicos
    data['Unique_Service'] = data['Familia'] + data['Tipo de Equipo'] + data['Tipo de Servicio']

    # Filtrar los datos por tienda
    filtered_df = data[data['Tienda'] == store_name].copy()

    # Convertir columnas de fecha a datetime
    for col in ['Ejec.1', 'Ult. Prev.']:
        filtered_df[col] = pd.to_datetime(filtered_df[col], format='%Y-%m-%d', errors='coerce')

    # Añadir verificación de columnas antes de continuar
    for col in ['Unique_Service', 'Ult. Prev.', 'Ejec.1']:
        if col not in filtered_df.columns:
            st.error(f"La columna {col} no se encuentra en los datos filtrados.")
            return

    # Obtener el mes actual
    current_month = datetime.now().month

    # Filtrar por los meses de enero al mes en curso y quedarse con la fecha más reciente en Ejec.1
    filtered_df_jan_to_now = filtered_df[(filtered_df['Ejec.1'].dt.month >= 1) & (filtered_df['Ejec.1'].dt.month <= current_month)]
    filtered_df_jan_to_now = filtered_df_jan_to_now.loc[filtered_df_jan_to_now.groupby('Unique_Service')['Ejec.1'].idxmax()]

    # Filtrar por los meses posteriores y quedarse con la fecha más reciente en Ult. Prev.
    filtered_df_next_months = filtered_df[(filtered_df['Ejec.1'].isna()) & (filtered_df['Ult. Prev.'].dt.month > current_month)]
    filtered_df_next_months = filtered_df_next_months.loc[filtered_df_next_months.groupby('Unique_Service')['Ult. Prev.'].idxmax()]

    # Combinar los dos dataframes
    final_df = pd.concat([filtered_df_jan_to_now, filtered_df_next_months])

    # Verificar y añadir servicios únicos en meses posteriores
    new_services = filtered_df[~filtered_df['Unique_Service'].isin(final_df['Unique_Service'])]
    if not new_services.empty:
        new_services = new_services.loc[new_services.groupby('Unique_Service')['Ult. Prev.'].idxmax()]
        final_df = pd.concat([final_df, new_services])

    # Seleccionar las columnas necesarias
    final_df = final_df[columns_to_include].copy()

    # Crear la nueva columna 'Ult. Preventivo'
    final_df['Ult. Preventivo'] = final_df['Ejec.1'].combine_first(final_df['Ult. Prev.'])

    # Añadir la columna Prog.1
    max_date = datetime.strptime('2024-12-31', '%Y-%m-%d')
    final_df['Prog.1'] = final_df['Ult. Preventivo'] + pd.to_timedelta(final_df['Frecuencia'], unit='D')
    final_df['Prog.1'] = final_df['Prog.1'].apply(lambda x: x if x <= max_date else None)

    # Formatear las fechas a DD-MM-YYYY
    for col in ['Ult. Prev.', 'Ejec.1', 'Ult. Preventivo', 'Prog.1']:
        final_df[col] = pd.to_datetime(final_df[col], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

    # Guardar los datos en un archivo Excel
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        worksheet_name = 'Fechas Planificadas'
        final_df.to_excel(writer, index=False, sheet_name=worksheet_name, startrow=2)
        worksheet = writer.sheets[worksheet_name]
        worksheet.write('A1', f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}')
        bold = writer.book.add_format({'bold': True})
        worksheet.set_row(0, None, bold)
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
        data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

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

    # Botón para generar el Excel filtrado
    if selected_store:
        if st.sidebar.button('Programa Anual de Mantenimiento'):
            if selected_month or selected_brand or selected_family:
                st.sidebar.warning("Por favor, deje solo el filtro de tienda lleno.")
            else:
                planned_excel_data = generate_excel(data, selected_store)
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
