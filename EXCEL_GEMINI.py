"""
EXCEL_GEMINI
"""

#!/usr/bin/env python
# coding: utf-8

# In[8]:


import os
import openpyxl

# Function to obtain creation date
def obtener_fecha_creacion(archivo):
    return os.path.getctime(archivo)

# Function to obtain last modification date
def obtener_fecha_modificacion(archivo):
    return os.path.getmtime(archivo)

# Function to handle potential missing author information
def obtener_autor_excel(archivo):
    try:
        libro_excel = openpyxl.load_workbook(archivo)  # Open in read-only mode (recommended)
        propiedades = libro_excel.properties
        # Access author using dictionary lookup (updated for compatibility)
        return propiedades.get("creator", "Desconocido")  # Use 'creator' if 'author' is unavailable
    except Exception as e:  # Catch potential errors gracefully
        print(f"Error retrieving author for '{archivo}': {e}")
        return "Error"  # Or a custom error value

# Function to extract information from Excel files
def extraer_informacion(ruta_carpeta):
    informacion = []
    archivos = os.listdir(ruta_carpeta)

    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        if archivo.endswith(".xlsx"):
            try:
                fecha_creacion = obtener_fecha_creacion(ruta_archivo)
                fecha_modificacion = obtener_fecha_modificacion(ruta_archivo)
                autor = obtener_autor_excel(ruta_archivo)
                informacion.append((archivo, fecha_creacion, fecha_modificacion, autor))
            except Exception as e:  # Handle errors for individual files
                print(f"Error processing '{archivo}': {e}")

    return informacion

# Function to save information to a new Excel file
def guardar_informacion(informacion, ruta_archivo_nuevo):
    libro_excel = openpyxl.Workbook()
    hoja_calculo = libro_excel.active

    hoja_calculo.append(["Nombre archivo", "Fecha creación", "Fecha modificación", "Autor"])
    for fila in informacion:
        hoja_calculo.append(fila)

    libro_excel.save(ruta_archivo_nuevo)
    print(f"Información guardada en: {ruta_archivo_nuevo}")

# Replace with your folder and output file paths
ruta_carpeta = "C:\\Users\\HP\\Downloads\\Tareas"
ruta_archivo_nuevo = "C:\\Users\\HP\\Downloads\\nuevo_archivo.xlsx"

informacion = extraer_informacion(ruta_carpeta)
guardar_informacion(informacion, ruta_archivo_nuevo)




# In[ ]:






if __name__ == "__main__":
    pass
