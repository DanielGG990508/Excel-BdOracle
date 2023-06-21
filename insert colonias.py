"""
Created on Mon May  8 17:16:43 2023

@author: 10042891 Daniel Garcia Garcia
Description: Este código utiliza las librerías "cx_Oracle" y "openpyxl" para insertar datos desde un archivo Excel a una base de datos Oracle. 
Primero se establece la conexión a la base de datos y se carga el archivo Excel. Luego se define la consulta de inserción y se selecciona la hoja del archivo Excel donde se encuentran los datos.
Después, se itera sobre las filas del archivo Excel y se extraen los datos de la columna A y B. Se ejecuta la consulta con los valores correspondientes de cada fila y se confirma la inserción al
 final del ciclo. Finalmente se cierra la conexión y se muestra el número de filas insertadas.
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
sql = "INSERT INTO TABLA(PARAMETROS) VALUES (:1, :2, :3, :4, :5,:6)"
# Seleccionar la hoja AGEEML
worksheet = workbook["Hoja1"]
iterador = 0
# Iterar sobre las filas y extraer los datos de la columna A y B
fecha_actual = datetime.datetime.now().strftime("%d/%m/%y")
for row in worksheet.iter_rows(min_row=2, values_only=True):
  dato_a = row[1]
  dato_b = row[2]
  dato_c = row[3]
  dato_d = row[4]
  iterador = iterador + 1
  #Se ejecuta la consulta con los valores correspondientes
  cursor.execute(sql, (dato_b,dato_a,fecha_actual,"USUARIO",dato_c,dato_d))
# Se confirma la inserción
conn.commit()
cursor.close()
conn.close()
print("conexion exitosa SE INSERTARON :",{iterador},"datos")

# Se cierra el cursor y la conexión a la base de datos
