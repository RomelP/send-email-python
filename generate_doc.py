#---------------------------------------------
# Este script realiza la colocación y envio 
# por email documento en formato pdf # 
# @utor: Romel Principe A # 
#---------------------------------------------
import os
#
from PIL import Image
from PIL import ImageDraw, ImageFont
#
import pandas as pd
import numpy as np
#---------------------------------------------
#    Cargar nombre de los participantes
#---------------------------------------------
# Cargando datos
df = pd.read_excel("Lista.xlsx", sheet_name="nombres", header= 0)   
#Visualizando los nombres
print("Apellidos y nombres :" )
print("#---------------------------------")
for i in range(134 , 159): #len(df)
     #nombre = df.iloc[i, 0]
     nombre = df["NOMBRES"][i]
     print(nombre)
# Separación 
print("#---------------------------------")
# Cantidad de filas
print("Número de certificados: ", len(df))
# Separación de impresion en pantalla
print("#---------------------------------")
#---------------------------------------------
# Genera las constancias cargando los datos 
#---------------------------------------------

# Cargando los datos apartir de la lista
hoja = pd.read_excel("Lista.xlsx", sheet_name="nombres",header= 0) 
# Generando las constancias de  manera recurrente
for i in range(134 , 159): #len(hoja)
    nombre = hoja["NOMBRES"][i] # lista de valores
    # Cargando la imagen 
    imagen = Image.open("Constancias.png")
    #
    draw = ImageDraw.Draw(imagen)
    # Tamaño y tipo de letra - nombre y codigo 
    font1 = ImageFont.truetype("/Library/Fonts/Arial Unicode.ttf", 90)
    # Cargar nombre 
    #draw.text((1060, 980), nombre, font=font1, fill="black") # nombres  # 0 - 134
    draw.text((920, 980), nombre, font=font1, fill="black") # nombres1 # 134 - 159
    #draw.text((810, 980), nombre, font=font1, fill="black") #  nombres2 # 159 - 161
    # Cargar codigo 
    imagen.save("certificados/" + str("constancia"+ str(i+1) + ".pdf")) # Path donde se guarda
    print("Certificado número:", i+1) 

# Separación de impresion en pantalla
print("#---------------------------------")