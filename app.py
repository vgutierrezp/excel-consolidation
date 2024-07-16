import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO
import datetime

# Function to add months with a limit date
def add_months_with_limit(source_date, months, max_date):
    try:
        new_date = source_date + pd.DateOffset(months=months)
        if new_date > max_date:
            return max_date
        return new_date
    except pd.errors.OutOfBoundsDatetime:
        return max_date

# Function to generate the plan
def generate_excel_with_dates(original_df, filtered_df, store_name):
    max_date = pd.Timestamp('2024-12-31')
    
    # Concatenate specific columns to identify unique services
    filtered_df['Unique_Service'] = filtered_df['Familia'] + filtered_df['Tipo de Equipo'] + filtered_df['Tipo de Servicio'] + filtered_df['Frecuencia'].astype(str)
    
    # Filter rows with non-empty 'Marca' column
    filtered_df = filtered_df[filtered_df['Marca'].notna()]
    
    # Group by 'Unique_Service' and get the most recent 'Ejec.1'
    new_df = filtered_df.loc[filtered_df.groupby('Unique_Service')['Ejec.1'].idxmax()]
    
    # Initialize the new DataFrame for the plan
    plan_columns = ['Tienda', 'Familia', 'Tipo de Equipo', 'Tipo de Servicio', 'Frecuencia', 'N° Equipos', 'Ult. Prev.']
    plan_df = new_df[plan_columns].copy()
    
    # Generate Prog.1, Prog.2, etc. columns ensuring the frequency interval
    for i in range(1, 29):
        plan_df[f'Prog.{i}'] = plan_df.apply(lambda row: add_months_with_limit(row['Ult. Prev.'], i * int(row['Frecuencia']), max_date), axis=1)
    
    # Ensure the difference between consecutive 'Prog.' columns is equal to the frequency
    for i in range(2, 29):
        plan_df[f'Prog.{i}'] = plan_df.apply(lambda row: add_months_with_limit(row[f'Prog.{i-1}'], int(row['Frecuencia']), max_date), axis=1)
    
    # Convert dates to required format
    for col in plan_df.columns:
        if 'Prog.' in col or col == 'Ult. Prev.':
            plan_df[col] = pd.to_datetime(plan_df[col]).dt.strftime('%d/%m/%Y')

    # Create the Excel writer and add the plan DataFrame
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    plan_df.to_excel(writer, sheet_name=store_name, index=False)
    
    # Add header with store name
    workbook = writer.book
    worksheet = writer.sheets[store_name]
    worksheet.merge_cells('A1:G1')
    worksheet.cell(row=1, column=1).value = f'PLAN ANUAL DE MANTENIMIENTO DE LA TIENDA: {store_name}'
    
    writer.save()
    processed_data = output.getvalue()
    return processed_data

# Main function to run the app
def main():
    st.title('Programa de Mantenimiento Preventivo')
    
    # Load data
    data_file = 'consolidated_file.xlsx'  # Update the path as needed
    data = pd.read_excel(data_file)
    
    # Handle months in Spanish and their order
    months = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
    }
    
    # Sidebar filters
    mes = st.sidebar.selectbox('Mes', [''] + sorted(data['Mes'].dropna().unique(), key=lambda x: months.get(x.lower(), 0)))
    marca = st.sidebar.selectbox('Marca', [''] + sorted(data['Marca'].dropna().unique()))
    tienda = st.sidebar.selectbox('Tienda', [''] + sorted(data['Tienda'].dropna().unique()))
    familia = st.sidebar.selectbox('Familia', [''] + sorted(data['Familia'].dropna().unique()))
    
    # Filter data based on selections
    filtered_data = data.copy()
    if mes:
        filtered_data = filtered_data[filtered_data['Mes'].str.lower() == mes.lower()]
    if marca:
        filtered_data = filtered_data[filtered_data['Marca'] == marca]
    if tienda:
        filtered_data = filtered_data[filtered_data['Tienda'] == tienda]
    if familia:
        filtered_data = filtered_data[filtered_data['Familia'] == familia]
    
    st.dataframe(filtered_data)
    
    # Input for store name
    store_name = st.text_input('Nombre de la Tienda', value='BB - Caminos Del Inca')
    
    # Button to generate the plan
    if st.button('Programa Anual de Mantenimiento'):
        planned_excel_data = generate_excel_with_dates(data, filtered_data, store_name)
        st.download_button(
            label=f'Descargar Plan Anual de Mantenimiento ({store_name})',
            data=planned_excel_data,
            file_name=f'Plan Anual de Mantenimiento {store_name}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == '__main__':
    main()
