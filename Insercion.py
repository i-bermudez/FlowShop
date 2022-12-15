import numpy as np
import openpyxl 
import random
from itertools import permutations
import math
import time

# Funcion para calcular makespan

def makespan(secuencia,tiempos):
    m = len(tiempos[0])  # Numero de maquinas
    n = len(tiempos)     # Numero de trabajos
    
    # Matriz organizada de acuerdo a la secuencia
    matriz = [[tiempos[i][j] for j in range(m)] for i in secuencia]
    
    make = [[0 for j in range(m)] for i in range(n)]
    
    make[0][0] = matriz[0][0]
    
    for i in range(1,n):
        for j in range(1,m):
            make[0][j] = make[0][j-1] + matriz[0][j]
            make[i][0] = make[i-1][0] + matriz[i][0]
    
    for i in range(1,n):
        for j in range(1,m):
            if make[i][j-1] > make[i-1][j]:
                make[i][j] = make[i][j-1] + matriz[i][j]
            else:
                make[i][j] = make[i-1][j] + matriz[i][j]
    
    return(make[-1][-1])

dataframe = openpyxl.load_workbook("InstanciasTaillard.xlsx")

# Define variable to read sheet
datos = dataframe["10"]

# Trabajos
n = 500

# Maquinas
m = 20

# Tiempos de procesamiento
t = []

for row in range(1,n+1):
    lista = []
    for col in range(1,m+1):
        lista.append(datos.cell(row,col).value)
    t.append(lista)

# ======================== BUSQUEDA LOCAL: INSERCIÓN ==========================

def insercion(lista, ins_pos1, en_pos2):
    
    x=lista.pop(ins_pos1)
    lista.insert(en_pos2,x)
    
    return lista

t_inicial = time.time()

mejor_solucion = list(range(n))
Cmax_best = makespan(mejor_solucion,t)
encontro = True

while encontro==True:
    encontro = False
    for i in range(10000):
        pos1 = random.randint(0,499)
        pos2 = random.randint(0,499)
        solucion = insercion(mejor_solucion,pos1,pos2)
        Cmax = makespan(solucion,t)
        if Cmax < Cmax_best:
            mejor_solucion = solucion[:]
            Cmax_best = Cmax
            encontro = True

for x in range(len(mejor_solucion)):
    mejor_solucion[x] = mejor_solucion[x]+1      

t_total = time.time()-t_inicial

print("Secuencia: ", mejor_solucion)
print("Cmax: ", Cmax_best)
#print(intercambio)
print(t_total/60)