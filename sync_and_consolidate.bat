@echo off
REM Cambiar a la ruta donde se encuentra el script consolidate.py
cd C:\Users\vgutierrez\chatbot_project

REM Ejecutar el script de consolidación
python consolidate.py

REM Copiar el archivo consolidado a la ubicación deseada en OneDrive (opcional)
copy consolidated_file.xlsx "C:\Users\vgutierrez\OneDrive - Servicios Compartidos de Restaurantes SAC\Documentos\01 Plan Preventivo Anual NGR\Preventivo\2024 PAM\PROVEEDORES"
