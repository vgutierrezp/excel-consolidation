import streamlit as st
import pandas as pd

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
    st.title("Visor de Datos Consolidado")

    data = load_data()

    if data.empty:
        st.error("No se pudieron cargar los datos.")
        return

    # Mostrar los datos
    st.write(data)

if __name__ == "__main__":
    main()
