# -*- coding: utf-8 -*-
"""
Created on Fri May 12 13:35:45 2023

@author: 10042891Daniel Garcia Garcia 

Description 
Este código carga un archivo Excel y extrae información de una columna específica. 
Luego utiliza expresiones regulares para extraer el nombre completo, números de teléfono y extensiones 
de la información extraída. Después crea un nuevo archivo Excel con la información extraída y los datos 
originales de la columna cargada en el archivo original. En resumen, este código procesa datos de un archivo Excel utilizando 
expresiones regulares para extraer información específica y guardarla en un nuevo archivo Excel.
"""
import pandas as pd
import re

# Cargar archivo Excel
df = pd.read_excel("ruta.xlsx", header=None)


# Expresiones regulares
patronN = r'expresion'
patronT = r'expresion'
patronE = r'expresion)'


#DICIDIR EL TEXTO EL PERDONAS 
personaSep=[]
# Extraer datos de la columna "Texto"
nombres = []
numeros = []
extensiones = []
numerosId = []
ids =[]
ides=0
for texto in df[4][1:]:
 # Encontrar todas las coincidencias
    coincidencias = re.findall(patronN, texto)
    
    # Dividir el texto en subelementos utilizando las coincidencias
    subelementos = re.split('|'.join(coincidencias), texto)
    
    # Crear la lista de sublistas
    sublistas = []
    for i in range(len(coincidencias)):
        sublistas.append([coincidencias[i], subelementos[i+1]])
    
    # Agregar la lista de sublistas a los resultados
    personaSep.append(sublistas)
#print(personaSep)
# Imprimir los resultados
for lista in personaSep:
    for elemento in lista:
        #print(str(elemento))
        matchN = re.search(patronN, str(elemento))
        matchT = re.search(patronT, str(elemento))
        matchE = re.search(patronE, str(texto))
        nombre = matchN.group().replace("Presidente", "*").replace("Presidenta", "*") if matchN else ""
        numero = matchT.group().replace(" ", "").replace("-", "") if matchT else ""
        extension = ", ".join(re.findall(r'\d{4}', matchE.group().replace("Red", ""))) if matchE else ""
        if len(numero) <= 5:
            numero = ""

        # Verificar si el número es igual a "304"
        if numero == "304":
            numero = ""
        ids.append(numerosId[ides])
        nombres.append(nombre)
        numeros.append(numero)
        extensiones.append(extension)
    ides+=1

#print(ids)
#Guardar cambios en el archivo Excel
df_nuevo = pd.DataFrame({"Id persona Principal":ids,"Nombre completo": nombres,"Numero":numeros,"Extension":extensiones})
df_nuevo.to_excel("RUTA", sheet_name="HOJA", index=False)
print("Se creó el archivo xlsx")
