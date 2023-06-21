# -- coding: utf-8 --
"""
Created on Tue May  2 17:48:14 2023

@author: 10042891
"""

# -- coding: utf-8 --
"""Copia de SCRAPING INFO.ipynb
# -- coding: utf-8 --
"""

import pandas as pd
#import requests
from bs4 import BeautifulSoup
#import regex as re
#import nltk
#from nltk.corpus import stopwords
#import matplotlib.pyplot as plt

archivo = open("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\circuito.txt","r")
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
    if (nueva_seccion and (child_a[i]['class'][-1] == expresiones[0] or child_a[i]['class'][-1] == expresiones[1])) or (i==len(child_a) - 1):      
      for j in range(len(r)):
          l = r[j]
          archivo.append([])
          archivo[x].append(valor_especial)
          for n in range(len(c)):
              archivo[x].append(c[n])
          archivo[x].append(l)
          archivo[x].append(persona[j])
          x += 1
      c = []
      r = []
      recorrerRed=0
      persona=[]
    if child_a[i]['class'][-1] == expresiones[0]:
        valor_especial = child_a[i].text
    elif child_a[i]['class'][-1] == expresiones[1]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == expresiones[2]:
        c.append(child_a[i].text)
    elif child_a[i]['class'][-1] == 'row':
        r.append(child_a[i].text)
        #al procesar un row, abrir el archivo de texto correspondiente y leer su contenido
        archivorRed = open("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\DatosRed\\circuito"+str(numLink)+".txt","r")
        datosRed = archivorRed.read()
        archivorRed.close()
        datosRed = BeautifulSoup(datosRed, 'html.parser')
        osoup1=datosRed
        # encontrar el div con la información deseada y agregarla a la lista 'archivo'
        childb = osoup1.find('div', {'class':'row alert-ojregistro'})
        if childb is not None:
            child_b= childb.find_all('div',{'style':'text-align:left'})
            texto = ''
            for elem in child_b:
               texto += elem.text+';'
            persona.append([])
            if recorrerRed < len(persona):
               persona[recorrerRed].append(texto)
            else:
               continue
        else:
           persona.append([])
           pass
       # incrementar el contador de archivos de texto
        recorrerRed += 1
        numLink+=1
    elif child_a[i]['class'][-1] == expresiones[3]:
        nueva_seccion = True

print("Archivo final creado")
print(f"tamaño {len(archivo)}")
archivo[0]
# Crear DataFrame desde la lista de listas 'archivo'
df = pd.DataFrame(archivo)
# Guardar DataFrame en un archivo Excel
df.to_excel("C:\\Users\\10042891\\.spyder-py3\\DatosRed\\PrimerCircuito\\DataDP2.xlsx", sheet_name="DataDP2", index=False)