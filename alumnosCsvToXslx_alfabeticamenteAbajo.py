import csv
from openpyxl import Workbook
import json

def setFila(contador_personas, fila):
    if contador_personas % 8 == 0:
        fila += 1

    fila += 1
    return fila

def switch_column(contador_personas, n_half_alu, fila, columna):
    if columna == 1 and contador_personas >= n_half_alu and contador_personas % 8 == 0:
        columna = 4
        fila = 1
        contador_personas = 0
    return {'contador_personas':contador_personas, 'columna':columna, 'fila':fila}

def paint_row(hoja: any, data_list : list, fila : int, columna : int, contador : int=0):
    hoja.cell(row=fila, column=columna).value = data_list[contador][0] # Nombre
    hoja.cell(row=fila, column=columna+1).value = data_list[contador][1] # Apellido
    return hoja

# Leer el archivo csv y guardarlo en una lista
nombres_y_apellidos = []
with open("asistentes.csv", newline="", encoding="utf-8") as file:
    reader = csv.reader(file)
    for row in reader:
        nombres_y_apellidos.append(row)

#print(json.dumps(nombres_y_apellidos, indent=4))

# Ordenar los alumnos por apellidos
nombres_y_apellidos = sorted(nombres_y_apellidos, key=lambda x: (x[1], x[0]))

# Creamos el archivo xlsx
workbook = Workbook()

# Seleccionamos la hoja activa
hoja = workbook.active

contador_personas = 1 # Contador para saber cuantos personas
fila = 1 # Fila en la que comienza (El xlsx comienza en 1)
columna = 1 # Columna en la que comienza (El xlsx comienza en 1)

# obtnemos el número de alumnos y la cantidad dividida entre dos para hacer 2 columnas
n_alu = len(nombres_y_apellidos)
n_half_alu = n_alu/2

# Iteramos sobre los nombres y apellidos
for i in range(n_alu):

    hoja = paint_row(hoja, nombres_y_apellidos, fila, columna, contador=i)

    fila = setFila(contador_personas, fila)
    
    result_column = switch_column(contador_personas,n_half_alu, fila, columna)
    columna = result_column.get('columna')
    fila = result_column.get('fila')
    contador_personas = result_column.get('contador_personas')

    contador_personas += 1

workbook.save("asistentes.xlsx")