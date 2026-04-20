import pandas as pd
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

#Se define la ruta del archivo donde se obtiene la info.
ruta = r"C:\Users\kavo6\Downloads\Base Inicial.xlsx"
df = pd.read_excel(ruta, sheet_name='Cálculos extra')

#Se establece qué variables dentro del dataframe se usarán para
#conformar el índice. En este caso mis dos variables fueron bautizadas con
#T1AP y T2AP
variables = df[['T1AP', 'T2AP']]

#En esta parte se estandarizan las variables para que le sea más fácil
#al algoritmo manejar los números, y el resultado no salga descompensado
#si algún vector tiene datos en diferente escala a los vectores. 
scaler = StandardScaler()
variables_scaled = scaler.fit_transform(variables)

#Se define el procedimiento de PCA, y se especifica que sólo se requiere
#una variable como componente final
pca = PCA(n_components=1)  
indice_compuesto = pca.fit_transform(variables_scaled)

#Se coloca el nuevo vector en el dataframe original
df['INDICE'] = indice_compuesto

#Se guarda un nuevo archivo excel con las modificaciones
df.to_excel(r"C:\Users\kavo6\Downloads\Base Final.xlsx", sheet_name='Sheet1', index=False)

