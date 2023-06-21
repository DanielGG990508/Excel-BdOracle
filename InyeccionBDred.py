# -*- coding: utf-8 -*-
"""
Created on Mon May 22 09:17:06 2023

@author: 10042891 Daniel Garcia Garcia
"""

from datetime import date
import pandas as pd
import requests,json
import re
###############################################################################################################################################################################################################
def obtenerRfc(nombre, apellido_paterno, apellido_materno):
    # Obtener la primera letra del apellido paterno
    letra_apellido_paterno = apellido_paterno[0] if apellido_paterno else ''

    # Obtener la primera letra del apellido materno o un 'X' si no se proporciona
    letra_apellido_materno = apellido_materno[0] if apellido_materno else 'X'

    # Obtener las primeras dos letras del nombre
    letras_nombre = nombre[:2] if nombre else ''

    # Concatenar las partes para formar el RFC
    rfc = letra_apellido_paterno + letra_apellido_materno + letras_nombre

    return rfc

def obtenerCurp(nombre, apellido_paterno, apellido_materno):
    # Obtener la primera letra del apellido paterno
    letra_apellido_paterno = apellido_paterno[0] if apellido_paterno else ''

    # Obtener la primera letra del apellido materno o un 'X' si no se proporciona
    letra_apellido_materno = apellido_materno[0] if apellido_materno else 'X'

    # Obtener la primera letra del nombre
    letra_nombre = nombre[0] if nombre else ''

    # Concatenar las partes para formar la CURP
    curp = letra_apellido_paterno + letra_apellido_materno + letra_nombre

    return curp
def formatear_nombre(nombre):
    palabras = nombre.split()
    palabras_formateadas = [palabra.capitalize() for palabra in palabras]
    return " ".join(palabras_formateadas)
###############################################################################################################################################################################################################

fechaP = date(1990, 1, 1)
fecha_formateada = fechaP.strftime("%d/%m/%Y")
# Cargar archivo Excel
df = pd.read_excel("RUTA", header=None)

#ides personas p
# Crear nuevas columnas para nombres y apellidos
df["Nombre"] = ""
df["Apellido Paterno"] = ""
df["Apellido Materno"] = ""

# Recorrer la columna de nombres completos
nombresL=[]
patronN=r'EXPRESION'

for nom in df[1][1:]:
    matchN = re.search(patronN, nom)
    if matchN is None:
        nombresC = nom.strip()
    else:
        nombresC = re.sub(patronN, '', nom).strip()
    nombresL.append(nombresC)
    #print(nombresC)
    #print('\n')
num_nombres = len(nombresL)
for i, texto in enumerate(nombresL):
    # Verificar si el nombre completo contiene palabras clave
    elementos = re.split(r';', str(texto.replace('*', '').replace('LIC. ', '').replace('MA.', '').replace('Lic', '')))
    print(elementos)
    listaN = ' '.join(elementos[0].split())
    nombres = listaN.split()
    num_nombres = len(nombres)
    # Verificar si el nombre completo contiene palabras clave
    palabras_clave = ["PALABRAS CLAVE]
    contiene_palabras_clave = any(palabra in texto.upper() for palabra in palabras_clave)
    if not contiene_palabras_clave:
        if num_nombres == 1:
            df.at[i, "Nombre"] = formatear_nombre(nombres[0])
        elif num_nombres == 2:
            df.at[i, "Nombre"] = formatear_nombre(nombres[0])
            df.at[i, "Apellido Paterno"] = formatear_nombre(nombres[1])
        elif num_nombres == 3:
            df.at[i, "Nombre"] = formatear_nombre(nombres[0])
            df.at[i, "Apellido Paterno"] = formatear_nombre(nombres[1])
            df.at[i, "Apellido Materno"] = formatear_nombre(nombres[2])
        elif num_nombres >= 4:
            if nombres[-1].lower() not in ["conmutador"]:
                df.at[i, "Nombre"] = formatear_nombre(' '.join(nombres[:-2]))
                df.at[i, "Apellido Paterno"] = formatear_nombre(nombres[-2])
                df.at[i, "Apellido Materno"] = formatear_nombre(nombres[-1])
            else:
                df.at[i, "Nombre"] = formatear_nombre(' '.join(nombres[:-3]))
                df.at[i, "Apellido Paterno"] =formatear_nombre( nombres[-3])
                df.at[i, "Apellido Materno"] =formatear_nombre( nombres[-2])
    else:
        secretario=''.join(nombres)
        df.at[i, "Nombre"] = secretario
#######################################################################################################################################
##edad y estado civil
    df.at[i,"Edad "]= 40
    df.at[i,"Estado civil"]= "Soltero"
########################################################################################################################################
##bloque de consulta para la api de genero 
    #nombre = nombres[0]  # Nombre para el cual deseas identificar el género
    #api_key = "mKhFG7tt2utQ5GvrFCXXDcaaVrgkt58b4ZT4"  # Reemplaza "tu_clave_de_api" con tu clave de la API
    # Realizar solicitud a la API
    #response = requests.get(f"https://api.genderize.io/?name="+ nombre).text
    # Analizar la respuesta JSON
    #genero = json.loads(response)['gender']
    df.at[i,"Genero"]= "F"
########################################################################################################################################
## fecha de nacimiento y nacionalidad         
    df.at[i,"Fecha Nacimiento "]= fecha_formateada 
    df.at[i,"Nacionalidad "]= "Mexico"
#######################################################################################################################################
##generacion de rfc y curp     
    if len(nombres) < 2 or nombres[-2] == "" or nombres[-1] == "":
        curp = "No se pudo generar CURP"
        rfc = "No se pudo generar RFC"
    else:
        curp = obtenerCurp(nombres[0], nombres[-2], nombres[-1])
        rfc = obtenerRfc(nombres[0], nombres[-2], nombres[-1])

########################################################################################################################################
##insercion de rfc y curp 
    df.at[i,"Curp "]= curp
    df.at[i,"Rfc"]= curp

##########################################################################################################################################
    #print(f"texto")
###########################################################################################################################################
##creacion de el documento excel
df = df.drop(columns=[1, 2, 3])  # Ajusta los números de las columnas según corresponda

    #DataFrame en un nuevo archivo Excel
df.to_excel("RUTA", sheet_name="RedInyeccion", index=False)

# Imprimir un mensaje indicando que se ha creado el archivo
print("Se creó el archivo xlsx")
