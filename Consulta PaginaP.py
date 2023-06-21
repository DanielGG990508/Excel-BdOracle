# -- coding: utf-8 --
"""
Created on Fri Apr 28 10:06:29 2023

@author: 10042891
"""



import pandas as pd
import requests
from bs4 import BeautifulSoup
import regex as re
import nltk
from nltk.corpus import stopwords
import matplotlib.pyplot as plt


archivo = open("RUTA DE ARCHIVO","r")
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
#print(f"Tamaño {len(expresiones)} Expresiones {expresiones}")
#print(expresiones)

len(osoup.find('div', {'class':'col-xs-12'}))

#nivel
#area
#ubicacion
#data = []
#for f in range(6):
  #print(osoup.findAll('div', {'class':'col-sm-4 alert-ojregistrotit'})[f].text, osoup.findAll('div', {'class':'col-sm-4 alert-ojregistro'})[f].text)
  #print(osoup.findAll('div', {'class':'col-sm-4 alert-ojregistrotit'})[f].text)

child = osoup.find('div', {'class':'col-xs-12'})
child_a = child.find_all('div',{'class':'row'})
#child_link = child.find_all('a',{'class':'fancybox'})

#for i in child_a:
#    print(i)


archivo = []
c = []
r = []
x = 0
l = ''
valor_especial = ""
# Se define una variable booleana que indica si se ha encontrado una nueva sección
nueva_seccion = False

for i in range(len(child_a)-1): 
    if child_a[i]['class'][-1] == expresiones[0]:
        valor_especial = child_a[i].text
    elif child_a[i]['class'][-1] == expresiones[1]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == expresiones[2]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == 'row':
        r.append(child_a[i].text)
    elif child_a[i]['class'][-1] == expresiones[3]:
            nueva_seccion = True
    if nueva_seccion and (child_a[i+1]['class'][-1] == expresiones[0] or child_a[i+1]['class'][-1] == expresiones[1]):
        for j in range(len(r)):
            l = r[j]
            archivo.append([])
            archivo[x].append(valor_especial)
            for n in range(len(c)):
                archivo[x].append(c[n])
            archivo[x].append(l)
            x += 1

        c = []
        r = []
#print (f"numero de linkss {len(child_links)}")
print("Archivo final creado")
print(f"tamaño {len(archivo)}")
#print(f" datos bonito \n{archivo[1]}\n")

#print(f" datos hardcore\n{archivo[220]}\n")
#archivo[0]



# Crear DataFrame desde la lista de listas 'archivo'
df = pd.DataFrame(archivo)

# Guardar DataFrame en un archivo Excel
df.to_excel("RUTA PARA GUARDAR RESULTADO", sheet_name="Data", index=False)
