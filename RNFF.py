#Se cargan las paqueterías necesarias
import pandas as pd
import tensorflow as tf
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

#Se carga la base de datos
ruta = r"C:\Users\kavo6\Downloads\Base Final.xlsx"
df = pd.read_excel(ruta, sheet_name='Sheet1')

#Se define la variable objetivo y las características ya elegidas
Y = df['APERSAUT']
X = df[['MRELOV', 'MBERHOOG', 'MINK3045', 'PBRAND', 'INDICE']]  

#Se dividen los datos en conjunto de entrenamiento y prueba
X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.2, random_state=42)

#Se reescalan los valores de los datos, lo que es un paso 
#fundamental, pues ayuda a que la red neuronal entrene de 
#manera más eficiente.
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

#Se crea el modelo de red neuronal Feedforward. Es usada una función de
#activación ReLU para las dos capas ocultas, y una sigmoide para la capa de salida. 
#La capas ocultas cuentan 64 y 32 neuronas por mera convención popular.
model = tf.keras.Sequential([
    tf.keras.layers.Dense(64, activation='relu', input_shape=(X_train_scaled.shape[1],)),
    tf.keras.layers.Dense(32, activation='relu'),  
    tf.keras.layers.Dense(1, activation='sigmoid')
])

#Se compila el modelo haciendo uso de un optimizador Adam para minimizar el
#function cost de entropía cruzada para medir qué tan bueno es el modelo.
model.compile(optimizer='adam',
              loss='binary_crossentropy',  
              metrics=['accuracy'])

#Se echa a andar el modelo
history = model.fit(X_train_scaled, Y_train, epochs=50, batch_size=32, validation_split=0.2, verbose=1)

#Se evalúa el modelo en el conjunto de prueba elegido
loss, accuracy = model.evaluate(X_test_scaled, Y_test)
print(f"Accuracy en el conjunto de prueba: {accuracy:.2f}")
