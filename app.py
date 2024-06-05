import streamlit as st
import pandas as pd
import os
from io import BytesIO

# Cargar los archivos Excel de la carpeta especificada
@st.cache_data
def load_data():
    # Asegúrate de que esta sea la ruta correcta en tu sistema
    folder_path = 'C:/Users/vgutierrez/OneDrive - Servicios Compartidos de Restaurantes SAC/Documentos/01 Plan Preventivo Anual NGR/Preventivo/2024 PAM/PROVEEDORES'
    
    # Comprobar si la ruta existe
    if not os.path.exists(folder_path):
        st.error(f"Directorio {folder_path} no encontrado.")
        return pd.DataFrame()  # Devolver un DataFrame vacío si la carpeta no existe
    
    all_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # Comprobar si hay archivos Excel en la carpeta
    if not all_files:
        st.error("No se encontraron archivos Excel en el directorio especificado.")
        return pd.DataFrame()  # Devolver un DataFrame vacío si no hay archivos

    # Cargar y concatenar todos los archivos de Excel en un solo DataFrame
    df_list = [pd.read_excel(os.path.join(folder_path, file)) for file in all_files]
    data = pd.concat(df_list, ignore_index=True)
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
    st.title("Navegador de Datos Consolidado")

    # Botón para actualizar los datos
    if st.sidebar.button('Actualizar Datos'):
        st.cache_data.clear()
    
    data = load_data()

    if data.empty:
        st.warning("No se pudieron cargar los datos. Verifica que los archivos Excel estén en la carpeta correcta.")
        return

    # Mostrar solo las columnas especificadas
    columns_to_show = ['Mes', 'Marca', 'Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    data = data[columns_to_show]

    # Formatear las columnas de fecha
    date_columns = ['Ult. Prev.', 'Prog.1', 'Ejec.1', 'CO', 'CL', 'IP', 'RP']
    for col in date_columns:
        data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d/%m/%y')

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
