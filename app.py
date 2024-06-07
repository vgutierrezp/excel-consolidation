import streamlit as st
import pandas as pd
from io import BytesIO

# URL del archivo en el repositorio de GitHub
url = 'https://raw.githubusercontent.com/vgutierrezp/excel-consolidation/main/consolidated_file.xlsx'

# Cargar los datos desde la URL
data = pd.read_excel(url)

# Función para ordenar los meses
def ordenar_meses(df, columna):
    meses_ordenados = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    df[columna] = pd.Categorical(df[columna], categories=meses_ordenados, ordered=True)
    return df

# Ordenar los meses en la columna 'Mes'
data = ordenar_meses(data, 'Mes')

# Título de la aplicación
st.title('Programa de Mantenimiento Preventivo')
st.write(f'Usando el archivo: {url}')

# Sidebar para filtros
st.sidebar.header('Filtros')
mes = st.sidebar.selectbox('Mes', data['Mes'].cat.categories)
marca = st.sidebar.selectbox('Marca', sorted(data['Marca'].unique()))
tienda = st.sidebar.selectbox('Tienda', sorted(data['Tienda'].unique()))
familia = st.sidebar.selectbox('Familia', sorted(data['Familia'].unique()))

# Aplicar filtros
filtered_data = data[(data['Mes'] == mes) & (data['Marca'] == marca) & (data['Tienda'] == tienda) & (data['Familia'] == familia)]

# Mostrar los datos filtrados
st.write(filtered_data)

# Botón para descargar los datos filtrados
def convertir_a_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.download_button(label='Descargar Excel', data=convertir_a_excel(filtered_data), file_name='filtered_data.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
