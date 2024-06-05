import streamlit as st
import pandas as pd
import os

# Función para cargar los datos
@st.cache_data
def load_data():
    try:
        # Ruta para el entorno local
        file_path_local = r'C:\Users\vgutierrez\chatbot_project\consolidated_file.xlsx'
        # Ruta para el entorno de despliegue
        file_path_deploy = 'consolidated_file.xlsx'

        if os.path.exists(file_path_local):
            file_path = file_path_local
        else:
            file_path = file_path_deploy

        st.write(f"Intentando cargar el archivo desde: {file_path}")  # Mensaje para depuración
        data = pd.read_excel(file_path)
        return data
    except FileNotFoundError as e:
        st.error(f"El archivo no se encontró: {e}")
    except Exception as e:
        st.error(f"Ocurrió un error al cargar los datos: {e}")
    return None

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
