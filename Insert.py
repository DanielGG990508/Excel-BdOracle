"""
Created on Mon May  8 17:16:43 2023

@author: 10042891 Daniel Garcia Garcia 

Description: 
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
sql = "INSERT INTO TABLA(PARAMETROS) VALUES (:1, :2, :3, :4, :5)"
# Seleccionar la hoja AGEEML
worksheet = workbook["AGEEML"]
iterador = 0
# Iterar sobre las filas y extraer los datos de la columna A y B
fecha_actual = datetime.datetime.now().strftime("%d/%m/%y")
for row in worksheet.iter_rows(min_row=2, values_only=True):
  dato_a = row[1]
  dato_b = row[2]
  dato_c = row[3]
  iterador = iterador + 1
  #Se ejecuta la consulta con los valores correspondientes
  cursor.execute(sql, (dato_c,fecha_actual,"Manuel",dato_a,dato_b))
# Se confirma la inserción
conn.commit()
cursor.close()
conn.close()
print("conexion exitosa SE INSERTARON :",{iterador},"datos")

