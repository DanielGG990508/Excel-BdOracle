import cx_Oracle
import openpyxl
import datetime
# Se establece la conexión a la base de datos
conn = cx_Oracle.connect('RVICEPRE/0rAcleDevVP2@10.204.14.120:1521')
# Definir la ruta y el nombre del archivo excel
ruta_excel = "C:/Users/10042891/Desktop/COColonia.xlsx"

# Cargar el archivo excel
workbook = openpyxl.load_workbook(ruta_excel)
# Se crea un cursor
cursor = conn.cursor()

# Se define la consulta de inserción
sql = "INSERT INTO TCVPCCOLONIAT2(FCCOLONIA,FCCP,FDFECHACT,FCUSERACT,FIIDCDFK,FIIDEDO) VALUES (:1, :2, :3, :4, :5,:6)"
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
  cursor.execute(sql, (dato_b,dato_a,fecha_actual,"Manuel",dato_c,dato_d))
# Se confirma la inserción
conn.commit()
cursor.close()
conn.close()
print("conexion exitosa SE INSERTARON :",{iterador},"datos")

# Se cierra el cursor y la conexión a la base de datos
