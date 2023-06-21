# -*- coding: utf-8 -*-
"""
Created on Wed May 22 17:45:26 2023

@author: 10042891
"""

import cx_Oracle
import openpyxl
import datetime
# Se establece la conexión a la base de datos
conn = cx_Oracle.connect('CREDENCIALES')
# Definir la ruta y el nombre del archivo excel
ruta_excel = "RUTA"

# Cargar el archivo excel
workbook = openpyxl.load_workbook(ruta_excel)
# Se crea un cursor
cursor = conn.cursor()

# Se define la consulta de inserción
sql = """INSERT INTO TABLA PARAMETROS)
         values( :1,:2, :3, :4, :5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15)"""# Seleccionar la hoja AGEEML
worksheet = workbook["RedInyeccion"]
iterador = 0
idsContacto=20
# Iterar sobre las filas y extraer los datos de la columna A y B
fecha_actual = datetime.datetime.now().strftime("%d/%m/%y")
for row in worksheet.iter_rows(min_row=2,values_only=True):
    PARAMETROSD[]
    cursor.execute(sql, (PARAMETROS))
    #este codigo es la insercion a la tabla de contacto que se hace simultanea 
    sql2 ="SELECT ELEMENTO FROM TABLA ORDER BY ID DESC FETCH FIRST 1 ROWS ONLY"
    # Ejecutar la consulta
    cursor.execute(sql2)
    resultado = cursor.fetchone()
    ultimo_dato = resultado[0] if resultado else None
    # Realizar la segunda inserción a otra tabla utilizando el ID
    sql3 = """INSERT INTO TABLA (PARAMETROS)
             values( :1,:2, :3, :4, :5)"""
  
    cursor.execute(sql3, (idsContacto,row[0],ultimo_dato,"ManuelRed",fecha_actual))
    idsContacto+=1
    print(f"{iterador} inserción realizada")

    # Se define la consulta de inserción
# Se confirma la inserción
conn.commit()
cursor.close()
conn.close()
print("conexion exitosa SE INSERTARON :",{iterador},"datos")

# Se cierra el cursor y la conexión a la base de datos
