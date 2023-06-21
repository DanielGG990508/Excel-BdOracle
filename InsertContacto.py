# -*- coding: utf-8 -*-
"""
Created on Tue May 23 18:35:37 2023

@author: 10042891
"""
import cx_Oracle
import openpyxl
import datetime
# Se establece la conexión a la base de datos
conn = cx_Oracle.connect('RVICEPRE/0rAcleDevVP2@10.204.14.120:1521')
# Definir la ruta y el nombre del archivo excel
ruta_excel = "C:/Users/10042891/.spyder-py3/DatosRed/PrimerCircuito/Match.xlsx"

# Cargar el archivo excel
workbook = openpyxl.load_workbook(ruta_excel)
# Se crea un cursor
cursor = conn.cursor()

# Se define la consulta de inserción
sql = """INSERT INTO TAVPCONTACTO (FIIDCONTACTO,FIIDPERSONAFK,FIIDOTRAPERSONA,FCUSER,FDFECHACT)
         values( :1,:2, :3, :4, :5)"""# Seleccionar la hoja AGEEML
worksheet = workbook["Sheet1"]
iterador =18
# Iterar sobre las filas y extraer los datos de la columna A y B
fecha_actual = datetime.datetime.now().strftime("%d/%m/%y")
for row in worksheet.iter_rows(min_row=2,max_row=113,values_only=True):
    personaP = row[0]
    personaC = row[2]
    iterador = iterador + 1
    cursor.execute(sql, (iterador,personaP,personaC,"DanielContacto",fecha_actual))
    print(f"se inserto el valor {iterador}")
# Se confirma la inserción
conn.commit()
cursor.close()
conn.close()
print("conexion exitosa SE INSERTARON :",{iterador},"datos")

# Se cierra el cursor y la conexión a la base de datos4
