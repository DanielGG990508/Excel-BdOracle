# -*- coding: utf-8 -*-
"""
Created on Thu May 25 16:36:54 2023
@author: 10042891 Daniel Garcia Garcia 
descripcion: Este código carga un archivo 
de Excel que contiene datos de texto en una 
columna y utiliza expresiones regulares para extraer 
información específica de cada registro.
"""

import pandas as pd
import re

# Cargar archivo Excel
df = pd.read_excel("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\DataDP2.xlsx", header=None)

#REVOLUCION NO. 1508, TORRE A, PISOS 2DO Y 3RO COL. GUADALUPE INN DELEG. ALVARO OBREGON CIUDAD DE MÉXICO 01020
# Definir expresiones regulares
patronCalle = r'(?i)(.*?)(?=\s+(NO\.|NÚMERO)|\b\d{1,4}\b)'
patronNC = r'(?i)\b\d{1,4}\b'
patronColonia = r'(?i)(COL\.\s+(.*?)\s)+(DELEG\.)|\b[A-ZÁ-Ú\s]+\b(?=, DELEGACIÓN|DELEGACION)'
#patronColonia = r'(?i)(COL\.\s+(.*?)\s)+(DELEG\.)|\b[A-ZÁ-Ú\s]+\b(?=, DELEGACIÓN|DELEGACION)|\d°\s(.*?)(\sMÉXICO|CIUDAD)'
patronCiudad = r'(?i)DELEG\.\s+(.*?)\s+(?=CIUDAD)'
patronCp = r'(?i)C\.P\.|\s+(\d{5})'
# Variables y listas para almacenar los resultados
calles = []
numeros = []
colonias = []
ciudades = []
codigoPs = []


# Recorrer cada texto en la columna "Texto"
for texto in df[2]:
    # Extraer nombre de la calle 
    matchCalle = re.match(patronCalle, str(texto))
    # Extraer número de calle
    matchNC = re.search(patronNC, str(texto))
    # Extraer colonia
    matchColonia = re.search(patronColonia, str(texto))
    # Extraer Ciudad
    matchCiudad = re.search(patronCiudad, str(texto))
    # Extraer CP
    matchCP = re.search(patronCp, str(texto))
########################################################################################### 
    #guardar Nombre de la calle 
    calle = matchCalle.group() if matchCalle else "N/A"
    # Extraer el número de CALLE
    numero = matchNC.group() if matchNC else "0000"
    # Extraer nombre de colonia 
    colonia = matchColonia.group().replace('COL.','').replace('DELEG.', '') if matchColonia else "N/A"
    # Extraer nombre de colonia 
    ciudad = matchCiudad.group().replace('DELEG.', '') if matchCiudad else "N/A"
    # Extraer nombre de colonia 
    codigoP = matchCP.group().replace('C.P.', "00000") if matchCP else "00000"
############################################################################################
    # Agregar los valores extraídos a las listas correspondientes
    calles.append(calle)
    numeros.append(numero)
    colonias.append(colonia)
    ciudades.append(ciudad)
    codigoPs.append(codigoP)
#############################################################################################
# Crear un nuevo DataFrame con los datos extraídos
df_nuevo = pd.DataFrame({"Direccion": df[2], "Calle": calles, "Números": numeros, "Colonia": colonias,"Ciudad":ciudades,"Codigo CP":codigoPs})

# Guardar el DataFrame en un nuevo archivo Excel
df_nuevo.to_excel("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\InyeccionDirecionesP.xlsx", sheet_name="InyeccionDirecionesP", index=False)

# Imprimir un mensaje indicando que se ha creado el archivo
print("Se creó el archivo xlsx")
