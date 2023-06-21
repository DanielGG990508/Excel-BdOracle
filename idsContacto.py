# -*- coding: utf-8 -*-
"""
Created on Tue May 23 12:24:59 2023

@author: 10042891
"""
import pandas as pd

# Leer el archivo de Excel
df = pd.read_excel("RUTA", sheet_name="Hoja1")

# Seleccionar la columna 'A' y 'B'
columna_a = df['Id persona Principal'] 

# Contar las ocurrencias de cada valor en la columna 'B' y ordenar los resultados
concurrecia = columna_a.value_counts().sort_index()
iterador=0
contador=0
idsN=[]
# Imprimir el valor y su recuento en orden ascendente
for valor, recuento in concurrecia.items():
    print(f"Valor: {valor}, Recuento: {recuento}")
    
    iterador+=1
    contador+=recuento
print(f"total de recuentos {contador}")
# Repetir los valores de la columna 'A' según el arreglo de repeticiones
#valores_repetidos = np.repeat(columna_a, repeticion)
#print(valores_repetidos)
print(f"total de Ids {len(idsN)}")
print(idsN)
df_repetidos = pd.DataFrame({"Valores repetidos": idsN})
# Guardar el DataFrame en un archivo Excel
print(f"total de dfrepetidos {len(df_repetidos)}")

print("Fin")
