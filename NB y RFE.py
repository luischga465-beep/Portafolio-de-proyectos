#Se importan las paqueterías necesarias para Naive Bayes y RFE. 
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.tree import DecisionTreeClassifier
from sklearn.feature_selection import RFE
from sklearn.metrics import accuracy_score

#Primero defino la ruta de mi archivo e indico la hoja a usar (Sheet1)
ruta = r"C:\Users\kavo6\Downloads\Base Final.xlsx"
df = pd.read_excel(ruta, sheet_name='Sheet1')

#Se explicita la variable objetivo, y las variables explicativas.
#En este caso, la matriz X es prácticamente todo el dataframe pero quitando
#tanto la variable objetivo, como otras columnas que no deberían usarse por su
#naturaleza tautológica, o cualitativa en lugar de cuantitativa. En general,
#para ello se usa el comando drop.
Y = df['APERSAUT']
X = df[['MRELOV', 'MBERHOOG', 'MINK3045', 'PBRAND', 'INDICE']]

#Se establece el conjunto de prueba y entrenamiento. Se propone que el conjunto
#de entrenamiento sea el 70% del total de datos. Asimismo, se coloca un randomstate
#para asegurar que se puedan obtener los mismos resultados al ejecutar
#posteriormente. 
X_train, X_test, y_train, y_test = train_test_split(X, Y, test_size=0.3, random_state=42)

#Se define el uso de un modelo decision tree, colocando un randomstate para que 
#no dé por resultado variables diferentes al correr después. 
tree_model = DecisionTreeClassifier(random_state=1)

#Aquí se define que quiero 5 variables a filtrar como resultado final
rfe = RFE(estimator=tree_model, n_features_to_select=5)

#Se define los datos que usaará el RFE, en este caso, el conjunto de entrenamiento
#que ya se había establecido antes.
rfe.fit(X_train, y_train)

#Se establece un nuevo conjunto de datos de entrenamiento y para prueba
#por medio de los resultados obtenidos en el RFE. 
X_train_rfe = rfe.transform(X_train)
X_test_rfe = rfe.transform(X_test)

#Se define el modelo Naive Bayes tipo Gaussiano, y se le alimenta con el conjunto
#de datos que se generó en el paso anterior con el RFE.
nb_model = GaussianNB()
nb_model.fit(X_train_rfe, y_train)

#Se calcula el modelo y se establece el nivel de precisión del mismo. 
y_pred = nb_model.predict(X_test_rfe)
accuracy = accuracy_score(y_test, y_pred)

print(f"Accuracy con las mejores características: {accuracy * 100:.2f}%")

#Se muestran las características seleccionadas
#X.columns[rfe.support_] signfica que se entra a la matriz original de X
#y se filtran las variables que hayan sido seleccionadas utilizando 
#la propiedad rfe.support_, guardando los nombres en selected_features. 
selected_features = X.columns[rfe.support_]
print(f"Características seleccionadas por RFE: {selected_features}")
