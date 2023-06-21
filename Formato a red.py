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
df = pd.read_excel("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\FormatoGeneralSoloNombres.xlsx", header=None)


# Expresiones regulares
patronN = r'(?!(?:COORDINADOR|SECRETARIO|SECRETARIA|ANALISTA|OFICIAL|PARTICULAR|COORDINATOR|MAGISTRADO|MAGISTRADA))\b(?:LIC\.|MA\.|[A-ZÁ-Úa-zá-ú]+\s[A-ZÁ-Úa-zá-ú]+\s[A-ZÁ-Úa-zá-ú\s]+)\s*(?:[A-ZÁ-Úa-zá-ú]+\.\s*)?[A-ZÁ-Úa-zá-ú\s]+'
patronT = r'\d(?:(?!\s{3,})[\d\s/-])*\d?'
patronE = r'(?:EXTS?|[Ee]xt)[\s:.,;]*(?:\d{4})(?:\s{0,7}[,Y]?\s*\d{1,4})*(?:\s*(?:Red|\b))|\d{1,4}(?:,\s*\d{1,4})*(?:\s{3,7}(?:\d{4},\s*)*\d{4}(?:Red))|\d{4}(?:Red)'


#DICIDIR EL TEXTO EL PERDONAS 
personaSep=[]
# Extraer datos de la columna "Texto"
nombres = []
numeros = []
extensiones = []
numerosId = [25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,
           71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,
           113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,
           148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,
           183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,
           218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,
           253,254,255,256,257,258,259,260,261,262,263,264,265,266,267,268,269,270,271,272,273,274,275,276,277,278,279,280,281,282,283,284,285,286,287,
           288,289,290,291,292,293,294,295,296,297,298,299,300,301,302,303,304,305,306,307,308,309,310,311,312,313,314,315,316,317,318,319,320,321,322,
           323,324,325,326,327,328,329,330,331,332,333,334,335,336,337,338,339,340,341,342,343,344,345,346,347,348,349,350,351,352,353,354,355,356,357,
           358,359,360,361,362,363,364,365,]
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
df_nuevo.to_excel("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\InyeccionDataRsoloN.xlsx", sheet_name="InyeccionDataRsoloN", index=False)
print("Se creó el archivo xlsx")
