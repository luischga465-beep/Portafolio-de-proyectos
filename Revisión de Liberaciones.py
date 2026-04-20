import pandas as pd

archivo_entrada = r"C:\Users\kavo6\Downloads\Extracción.xlsx"
archivo_salida = 'Revisión de Liberaciones.xlsx'

# Lista de nombres de las hojas que deseas procesar
hojas_a_procesar = ['SSC', 'Operador']

# Inicializar un diccionario para almacenar los datos procesados de cada hoja
datos_procesados = {}

# Función para convertir celdas de la columna 2 a número y luego a texto
def convertir_columna_2_a_numero_y_texto(df):
    df.iloc[:, 1] = df.iloc[:, 1].apply(lambda x: str(pd.to_numeric(x, errors='ignore')).replace(',', '').replace(' ', ''))
    return df.applymap(lambda x: str(x).replace(',', '').replace(' ', ''))

# Iterar sobre cada hoja y realizar el procesamiento
for hoja_nombre in hojas_a_procesar:
    # Lee los datos de la hoja actual
    datos = pd.read_excel(archivo_entrada, sheet_name=hoja_nombre, dtype=str)
    
    # Convierte la columna 2 a número y luego todo a tipo texto, eliminando comas y vacíos
    datos = convertir_columna_2_a_numero_y_texto(datos)
    
    # Almacena los datos procesados en el diccionario
    datos_procesados[hoja_nombre] = datos

# Convierte las dos columnas en conjuntos de tuplas (parejas)
parejas_hoja1 = set(datos_procesados['SSC'].apply(tuple, axis=1))
parejas_hoja2 = set(datos_procesados['Operador'].apply(tuple, axis=1))

# Encuentra las parejas que no están en la otra hoja
no_en_hoja2 = parejas_hoja1 - parejas_hoja2
no_en_hoja1 = parejas_hoja2 - parejas_hoja1

# Convierte los conjuntos en dataframes para escribirlos en el archivo Excel
no_en_hoja2_df = pd.DataFrame(list(no_en_hoja2), columns=datos_procesados['SSC'].columns)
no_en_hoja1_df = pd.DataFrame(list(no_en_hoja1), columns=datos_procesados['Operador'].columns)

# Buscar duplicados en la columna "Infracción" para cada hoja y añadir columna de origen
duplicados_hoja1 = datos_procesados['SSC'][datos_procesados['SSC'].duplicated(subset=['Infracción'], keep=False)].copy()
duplicados_hoja1['Hoja_Origen'] = 'SSC'

duplicados_hoja2 = datos_procesados['Operador'][datos_procesados['Operador'].duplicated(subset=['Infracción'], keep=False)].copy()
duplicados_hoja2['Hoja_Origen'] = 'Operador'

# Combina los duplicados de ambas hojas en un solo DataFrame
duplicados_df = pd.concat([duplicados_hoja1, duplicados_hoja2]).drop_duplicates()

# Calcular totales
totales_df = pd.DataFrame({
    'Hoja': ['SSC', 'Operador'],
    'Total de Parejas': [len(parejas_hoja1), len(parejas_hoja2)]
})

# Escribe los datos procesados en un nuevo archivo Excel
with pd.ExcelWriter(archivo_salida) as writer:
    no_en_hoja2_df.to_excel(writer, sheet_name='No encontrado en Operador', index=False)
    no_en_hoja1_df.to_excel(writer, sheet_name='No encontrado en SSC', index=False)
    duplicados_df.to_excel(writer, sheet_name='Infracciones Duplicadas', index=False)
    totales_df.to_excel(writer, sheet_name='Totales', index=False)
