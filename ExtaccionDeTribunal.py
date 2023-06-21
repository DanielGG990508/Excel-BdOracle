
import pandas as pd
#import requests
from bs4 import BeautifulSoup
#import regex as re
#import nltk
#from nltk.corpus import stopwords
#import matplotlib.pyplot as plt
archivo = open("RUTA DEL ARCHIVO ","r")
pagina = archivo.read()
archivo.close()
pagina = BeautifulSoup(pagina, 'html.parser')
osoup = pagina
recorrerRed=0
expresiones = [
    'alert-ojlevel',
    'alert-ojarea',
    'alert-ojubica',
    'alert-ojregistro',
]
child = osoup.find('div', {'class':'col-xs-12'})
child_a = child.find_all('div',{'class':'row'})
dataRed=[]
archivo = []
c = []
r = []
x = 0
l = ''
persona=[]
numLink=0
valor_especial = ""
nueva_seccion = False
for i in range(len(child_a)): 
#for i in range(10): 
    if child_a[i]['class'][-1] == expresiones[0]:
        valor_especial = child_a[i].text
    elif child_a[i]['class'][-1] == expresiones[1]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == expresiones[2]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == 'row':
        r.append(child_a[i].text)
        #al procesar un row, abrir el archivo de texto correspondiente y leer su contenido
        archivorRed = open("RUTA"+str(numLink)+".txt","r")
        datosRed = archivorRed.read()
        archivorRed.close()
        datosRed = BeautifulSoup(datosRed, 'html.parser')
        osoup1=datosRed
        #print(osoup1)
        # encontrar el div con la información deseada y agregarla a la lista 'archivo'
        childb = osoup1.find('div', {'class':'row alert-ojregistro'})
        #print(childb)
        if childb is not None:
            child_b= childb.find_all('div',{'style':'text-align:left'})
       
            #print(child_b)
            #print(len(child_b))
        ####################################################################################################################################
            texto = ''
            for elem in child_b:
                texto += elem.text
                #print(texto)
                #print(recorrerRed)
            persona.append([])
            if recorrerRed < len(persona):
                persona[recorrerRed].append(texto)
            else:
                continue
        else:
            persona.append([])
            pass
        recorrerRed += 1
        numLink+=1
    elif child_a[i]['class'][-1] == expresiones[3]:
            nueva_seccion = True
    if nueva_seccion and (child_a[i]['class'][-1] == expresiones[0] or child_a[i]['class'][-1] == expresiones[1]):
        for j in range(len(r)):
            l = r[j]
            archivo.append([])
            archivo[x].append(valor_especial)
            for n in range(len(c)):
                archivo[x].append(c[n])
            archivo[x].append(l)
            archivo[x].append(persona[j])
            x += 1 
        recorrerRed=0
        persona=[]
        c = []
        r = []
#print (f"numero de linkss {len(child_links)}")
print("Archivo final creado")
print(f"tamaño {len(archivo)}")
#print(f" datos bonito \n{archivo[1]}\n")
#print(f" datos hardcore\n{archivo[220]}\n")
archivo[0]
# Crear DataFrame desde la lista de listas 'archivo'
df = pd.DataFrame(archivo)
# Guardar DataFrame en un archivo Excel
df.to_excel("RUTA", sheet_name="DataD", index=False)
