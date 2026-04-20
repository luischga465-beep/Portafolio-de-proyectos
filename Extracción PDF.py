import os
import pdfplumber
import pandas as pd

carpeta_path = r"C:\Users\kavo6\Downloads\2. Lomas Chapultepec\OFICIOS DE LIBERACION JUN-2024-20240716T021146Z-001\OFICIOS DE LIBERACION JUN-2024"

def extraer_palabra_siguiente_a_palabra_clave(pdf_path, palabras_clave):
    palabras_extraidas = {palabra_clave: ":" for palabra_clave in palabras_clave}

    with pdfplumber.open(pdf_path) as pdf:
        for pagina_num in range(len(pdf.pages)):
            pagina = pdf.pages[pagina_num]
            texto_pagina = pagina.extract_text()

            for palabra_clave in palabras_clave:
                indice = texto_pagina.find(palabra_clave)
                if indice != -1 and indice + len(palabra_clave) < len(texto_pagina):
                    palabra_siguiente = texto_pagina[indice + len(palabra_clave):].split()[0]
                    palabras_extraidas[palabra_clave] = palabra_siguiente

    return palabras_extraidas

def procesar_archivos_en_carpeta(carpeta_path, palabras_clave):
    resultados = []

    for nombre_archivo in os.listdir(carpeta_path):
        if nombre_archivo.endswith('.pdf'):
            pdf_path = os.path.join(carpeta_path, nombre_archivo)
            palabras_extraidas = extraer_palabra_siguiente_a_palabra_clave(pdf_path, palabras_clave)
            resultados.append(palabras_extraidas)

    return resultados

# Ejemplo de uso
palabras_clave = ['placa:', 'placa', 'infracción:', 'instruye:']
resultados = procesar_archivos_en_carpeta(carpeta_path, palabras_clave)

# Crear un DataFrame con los resultados
df = pd.DataFrame(resultados)

# Copiar el contenido de la columna 'placa' a 'placa:' si 'placa:' contiene solo ':'
df.loc[df['placa:'] == ':', 'placa:'] = df['placa']

# Eliminar las columnas 'placa' y 'instruye:'
df = df.drop(columns=['placa', 'instruye:'])

# Renombrar las columnas 'placa:' a 'Placa' y 'infracción:' a 'Infraccion'
df = df.rename(columns={'placa:': 'Placa', 'infracción:': 'Infracción'})

# Exportar el DataFrame a un archivo Excel con la hoja llamada 'SSC'
archivo_excel = 'Extracción.xlsx'
with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='SSC')
    
    # Añadir una nueva hoja llamada "Operador"
    workbook  = writer.book
    worksheet = workbook.add_worksheet('Operador')
    
    # Escribir "Placa" y "Infracción" en las primeras celdas
    worksheet.write(0, 0, 'Placa')
    worksheet.write(0, 1, 'Infracción')
