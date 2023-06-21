# -*- coding: utf-8 -*-
"""
Created on Tue May  9 10:09:45 2023

@author: 10042891 Daniel Garcia Garcia 
descripcion: Este código carga un archivo 
de Excel que contiene datos de texto en una 
columna y utiliza expresiones regulares para extraer 
información específica de cada registro. Luego, 
los datos extraídos se guardan en un nuevo archivo de Excel.
 En concreto, el código busca el nombre completo de la persona mencionada en el registro,
 su número de teléfono y su(s) extensión(es), en caso de que se mencionen. El código utiliza 
 expresiones regulares para buscar patrones específicos en el texto y extraer la información deseada. 
 Los resultados se guardan en un nuevo archivo de Excel con las columnas "Texto", "Nombre completo", "Números" y "Extensión".
"""

import pandas as pd
import re

# Cargar archivo Excel
df = pd.read_excel("RUTA", header=None)

# Definir expresiones regulares
patronN = r'EXPRESION'
patronT = r'EXPRESION'
patronE = r'EXPRESION'

# Variables y listas para almacenar los resultados
nombre = ''
nombres = []
numeros = []
extensiones = []

# Recorrer cada texto en la columna "Texto"
for texto in df[3]:
    # Extraer nombres y apellidos
    matchN = re.match(patronN, str(texto))
    # Extraer números de teléfono
    matchT = re.search(patronT, str(texto))
    # Extraer extensiones telefónicas
    matchE = re.search(patronE, str(texto))
    
    # Extraer el nombre completo
    nombre = matchN.group().replace("Presidente", "*").replace("Presidenta", "*") if matchN else ""
    
    # Extraer el número de teléfono
    numero = matchT.group().replace(" ", "").replace("-", "") if matchT else ""
    
    # Extraer las extensiones telefónicas
    extension = ", ".join(re.findall(r'\d{4}', matchE.group().replace("Red", ""))) if matchE else ""
    
    # Agregar los valores extraídos a las listas correspondientes
    nombres.append(nombre)
    numeros.append(numero)
    extensiones.append(extension)

# Crear un nuevo DataFrame con los datos extraídos
df_nuevo = pd.DataFrame({"columna Principal": df[3], "Nombre completo": nombres, "Números": numeros, "Extensión": extensiones})

# Guardar el DataFrame en un nuevo archivo Excel
df_nuevo.to_excel("RUTA", sheet_name="HOJA", index=False)

# Imprimir un mensaje indicando que se ha creado el archivo
print("Se creó el archivo xlsx")
