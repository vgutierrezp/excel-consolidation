import streamlit as st
import pandas as pd

# Función para cargar los datos
@st.cache_data
def load_data():
    file_path = r'C:\Users\vgutierrez\chatbot_project\consolidated_file.xlsx'
    data = pd.read_excel(file_path)
    return data

# Función principal
def main():
    st.title("Visor de Datos Consolidado")

    data = load_data()

    if data is not None:
        st.write(data)
    else:
        st.error("No se pudieron cargar los datos.")

if __name__ == "__main__":
    main()
