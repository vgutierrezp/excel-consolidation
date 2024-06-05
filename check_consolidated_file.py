import pandas as pd

# Cargar el archivo consolidado
file_path = 'C:/Users/vgutierrez/chatbot_project/consolidated_file.xlsx'
data = pd.read_excel(file_path)

# Mostrar información básica del DataFrame
print("Información del DataFrame:")
print(data.info())

# Mostrar las primeras filas del DataFrame
print("\nPrimeras filas del DataFrame:")
print(data.head())

# Mostrar un resumen de las filas que contienen 'consolidated_file.xlsx' en la columna 'Source File'
filtered_data = data[data['Source File'] == 'consolidated_file.xlsx']
print("\nFilas que contienen 'consolidated_file.xlsx' en 'Source File':")
print(filtered_data)
