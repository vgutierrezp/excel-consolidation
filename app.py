import streamlit as st
import pandas as pd
from io import BytesIO

# Cargar el archivo consolidado desde el repositorio
@st.cache_data
def load_data():
    file_path = 'consolidated_file.xlsx'
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return pd.DataFrame()

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
    months = sorted([''] + list(data['Mes'].dropna().unique()))
    brands = sorted([''] + list(data['Marca'].dropna().unique()))
    stores = sorted([''] + list(data['Tienda'].dropna().unique()))
    families = sorted([''] + list(data['Familia'].dropna().unique()))

    # Crear filtros dependientes
    selected_month = st.sidebar.selectbox('Mes', options=months)
    filtered_data = data if selected_month == '' else data[data['Mes'] == selected_month]

    selected_brand = st.sidebar.selectbox('Marca', options=sorted([''] + list(filtered_data['Marca'].dropna().unique())))
    filtered_data = filtered_data if selected_brand == '' else filtered_data[filtered_data['Marca'] == selected_brand]

    selected_store = st.sidebar.selectbox('Tienda', options=sorted([''] + list(filtered_data['Tienda'].dropna().unique())))
    filtered_data = filtered_data if selected_store == '' else filtered_data[filtered_data['Tienda'] == selected_store]

    selected_family = st.sidebar.selectbox('Familia', options=sorted([''] + list(filtered_data['Familia'].dropna().unique())))
    filtered_data = filtered_data if selected_family == '' else filtered_data[filtered_data['Familia'] == selected_family]

    # Mostrar los datos filtrados
    st.write(filtered_data)

    # Opción para descargar el archivo filtrado
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data

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
