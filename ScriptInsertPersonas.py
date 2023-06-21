# -*- coding: utf-8 -*-
"""
Created on Wed May 22 17:45:26 2023

@author: 10042891
"""

import cx_Oracle
import openpyxl
import datetime
# Se establece la conexión a la base de datos
conn = cx_Oracle.connect('RVICEPRE/0rAcleDevVP2@10.204.14.120:1521')
# Definir la ruta y el nombre del archivo excel
ruta_excel = "C:/Users/10042891/.spyder-py3/DatosRed/PrimerCircuito/RedInyeccion.xlsx"

# Cargar el archivo excel
workbook = openpyxl.load_workbook(ruta_excel)
# Se crea un cursor
cursor = conn.cursor()

# Se define la consulta de inserción
sql = """INSERT INTO TAVPPERSONA (FBFOTO,FCNOMBRE,FCAPELLIDOP,FCAPELLIDOM,FIEDAD,FCESTADOCIVIL,
         FCGENERO,FDFECHANACI,FCNACIONALIDAD,FCCURP,FCRFC,FISTATUS,FCUSER,FDFECHACT,FDREG)
         values( :1,:2, :3, :4, :5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15)"""# Seleccionar la hoja AGEEML
worksheet = workbook["RedInyeccion"]
iterador = 0
idsContacto=20
# Iterar sobre las filas y extraer los datos de la columna A y B
fecha_actual = datetime.datetime.now().strftime("%d/%m/%y")
for row in worksheet.iter_rows(min_row=2,values_only=True):
    nombre = row[1]
    apellidoP = row[2]
    apellidoM = row[3]
    edad = row[4]
    estadoC = row[5]
    genero = row[6]
    fechaNaci=row[7]
    nacionalidad=row[8]
    curp=row[9]
    rfc=row[10]
    iterador = iterador + 1
    cursor.execute(sql, ("",nombre,apellidoP,apellidoM,edad,estadoC,genero,fechaNaci,nacionalidad,curp,rfc,1,"ManuelRed","",fecha_actual))
    #este codigo es la insercion a la tabla de contacto que se hace simultanea 
    sql2 ="SELECT fiidpersona FROM tavppersona ORDER BY fiidpersona DESC FETCH FIRST 1 ROWS ONLY"
    # Ejecutar la consulta
    cursor.execute(sql2)
    resultado = cursor.fetchone()
    ultimo_dato = resultado[0] if resultado else None
    # Realizar la segunda inserción a otra tabla utilizando el ID
    sql3 = """INSERT INTO TAVPCONTACTO (FIIDCONTACTO,FIIDPERSONAFK,FIIDOTRAPERSONA,FCUSER,FDFECHACT)
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
