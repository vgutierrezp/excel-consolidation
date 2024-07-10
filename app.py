import streamlit as st
import pandas as pd

# Definir la función principal
def main():
    st.title('Consolidación de Archivos Excel')

    # Cargar los datos
    data_file = 'consolidated_file.xlsx'
    data = pd.read_excel(data_file)

    # Definir los filtros
    st.sidebar.header('Filtros')
    mes = st.sidebar.selectbox('Mes', [''] + sorted(data['Mes'].dropna().unique().tolist()))
    marca = st.sidebar.selectbox('Marca', [''] + sorted(data['Marca'].dropna().unique().tolist()))
    tienda = st.sidebar.selectbox('Tienda', [''] + sorted(data['Tienda'].dropna().unique().tolist()))
    familia = st.sidebar.selectbox('Familia', [''] + sorted(data['Familia'].dropna().unique().tolist()))

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

# Ejecutar la aplicación
if __name__ == "__main__":
    main()
