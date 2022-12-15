import numpy as np
import openpyxl 
import random
from itertools import permutations
import math
import time
import matplotlib.pyplot as plt

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

# ======================== BUSQUEDA LOCAL: SWAP ==========================

def swapPosiciones(lista, pos1, pos2):
     
    lista[pos1], lista[pos2] = lista[pos2], lista[pos1]
    return lista

t_inicial = time.time()

mejor_solucion = list(range(n))
Cmax_best = makespan(mejor_solucion,t)
encontro = True

iteraciones = 0
itera = 0
x = []
y=[]

while iteraciones < 10000:
    posicionUno = random.randint(0,499)
    posicionDos = random.randint(0, 499)  
    solucion = swapPosiciones(mejor_solucion,posicionUno,posicionDos)
    Cmax = makespan(solucion,t)
    iteraciones +=1
    itera+=1
    if Cmax < Cmax_best:
        iteraciones =0
        itera+=1
        mejor_solucion = solucion[:]
        Cmax_best = Cmax
    x.append(itera)
    y.append(Cmax_best)

plt.plot(x,y,":")
plt.show()

for x in range(len(mejor_solucion)):
    mejor_solucion[x] = mejor_solucion[x]+1            

t_total = time.time()-t_inicial

print("Secuencia: ", mejor_solucion)
print("Cmax: ", Cmax_best)
print("Tiempo total: ", t_total/60)
    

    
    
    
    
    
