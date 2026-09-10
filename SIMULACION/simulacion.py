# importar archivo excel, para leer lo que hay dentro y que me muestre en una tabla (con print)
import pandas as pd
df = pd.read_excel(r"C:\Users\jenss\Desktop\Pyton\Proyecto 1\SIMULACION\simulacion.xlsx")
print(df)

# misma tabla pero para lo que hay dentro de la carpeta SIMULACION, que tiene los 4 municipios y el excel
import os
carpeta_base = r"C:\Users\jenss\Desktop\Pyton\Proyecto 1\SIMULACION"
#creo la variable, que toma los 4 valores de los 4 municipios
municipios = ["azul", "berisso", "laplata", "laprida", "merlo", "pilar"]

# hago bucle para:
#       1) Lee el contenido del excel
#       2) Compara los "ok", "no hay ordenanza" y "no hay informacion" del 
#   excel, con que haya o no archivo en la carpeta del respectivo municipio

for municipio in municipios:
    ruta_municipio = carpeta_base + "/" + municipio
    contenido = os.listdir(ruta_municipio)

    fila = df[df["municipio"] == municipio]
    valor_check = fila["check"].values[0]

    if contenido:
        # hay archivo en la carpeta
        if valor_check in ["ok", "no hay informacion"]:
            print(municipio, "-> CORRECTO")
        else:
            print(municipio, "-> INCONSISTENCIA. Excel dice:", valor_check, "| Carpeta: tiene archivo")
    else:
        # la carpeta está vacía
        if valor_check == "no hay ordenanza":
            print(municipio, "-> CORRECTO")
        else:
            print(municipio, "-> INCONSISTENCIA. Excel dice:", valor_check, "| Carpeta: vacía")