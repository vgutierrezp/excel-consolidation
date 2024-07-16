import pandas as pd
import numpy as np
import streamlit as st

# Función para obtener la fecha más reciente de 'Ejec.1'
def get_recent_execution_date(group):
    return group.loc[group['Ejec.1'].idxmax()]

# Función para obtener la fecha más reciente de 'Ult. Prev.' para nuevas concatenaciones
def get_recent_prev_date(group):
    return group.loc[group['Ult. Prev.'].idxmax()]

# Función principal
def main():
    st.title("Programa de Mantenimiento Preventivo")

    # Cargar los datos
    data = pd.read_excel("path/to/your/datafile.xlsx")

    # Asegurarse de que las fechas estén en el formato correcto
    data['Ult. Prev.'] = pd.to_datetime(data['Ult. Prev.'], format='%Y-%m-%d %H:%M:%S')
    data['Ejec.1'] = pd.to_datetime(data['Ejec.1'], format='%Y-%m-%d %H:%M:%S')

    # Filtros de la tienda
    selected_store = st.sidebar.selectbox('Tienda', data['Tienda'].unique())
    filtered_data = data[data['Tienda'] == selected_store]

    # Crear columna 'Unique_Service'
    filtered_data['Unique_Service'] = (filtered_data['Familia'] + 
                                       filtered_data['Tipo de Equipo'] + 
                                       filtered_data['Tipo de Servicio'])

    # Filtrar datos de Enero al mes en curso
    current_month = pd.to_datetime('now').month
    current_year = pd.to_datetime('now').year
    filtered_data['Mes_Num'] = pd.to_datetime(filtered_data['Ult. Prev.']).dt.month
    filtered_data['Year'] = pd.to_datetime(filtered_data['Ult. Prev.']).dt.year
    jan_to_current = filtered_data[(filtered_data['Mes_Num'] >= 1) & (filtered_data['Mes_Num'] <= current_month) & (filtered_data['Year'] == current_year)]

    # Obtener las filas con las fechas más recientes en 'Ejec.1'
    recent_executions = jan_to_current.groupby('Unique_Service').apply(get_recent_execution_date).reset_index(drop=True)

    # Filtrar datos del mes siguiente hasta Diciembre
    next_month_to_dec = filtered_data[(filtered_data['Mes_Num'] > current_month) & (filtered_data['Year'] == current_year)]

    # Obtener las filas con las fechas más recientes en 'Ult. Prev.'
    recent_prevs = next_month_to_dec.groupby('Unique_Service').apply(get_recent_prev_date).reset_index(drop=True)

    # Concatenar los resultados
    final_data = pd.concat([recent_executions, recent_prevs]).drop_duplicates(subset='Unique_Service')

    # Seleccionar las columnas necesarias
    final_data = final_data[['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Ejecutor', 'Frecuencia', 'N° Equipos', 'Ult. Prev.', 'Ejec.1']]

    # Guardar el resultado en un nuevo archivo Excel
    output_filename = f"Plan de Mantenimiento Anual - {selected_store}.xlsx"
    final_data.to_excel(output_filename, index=False)

    st.success(f"Archivo guardado como {output_filename}")

if __name__ == "__main__":
    main()
