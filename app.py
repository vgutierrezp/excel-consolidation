import streamlit as st
import pandas as pd
from io import BytesIO

# Cargar el archivo consolidado
@st.cache_data
def load_data():
    try:
        data = pd.read_excel('consolidated_file.xlsx')
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

# Función principal
def main():
    st.title("PLAN ANUAL DE MANTENIMIENTO PREVENTIVO")

    data = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Mostrar todos los datos sin filtrar las columnas
    st.write(data)

    # Verificar si hay duplicados y eliminarlos
    data = data.drop_duplicates()

    # Opción para descargar el archivo completo
    st.sidebar.header('Descargar Datos')
    if not data.empty:
        excel_data = to_excel(data)
        st.sidebar.download_button(
            label='Descargar Excel',
            data=excel_data,
            file_name='complete_data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
