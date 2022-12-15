import numpy as np
import openpyxl 
import random
from itertools import permutations
import math
import time
import matplotlib.pyplot as plt
import itertools

# ==================================================
# IMPORTAR DATOS DE INSTANCIAS TAILLARD
# ==================================================

# Funcion aplanar lista

def flatter(lst):
    ret = []
    for elem in lst:
        if isinstance(elem, list):
            ret.extend(flatter(elem))
        else:
            ret.append(elem)
    return ret

flatter([[1,3,2],12,3,[2,3]])

# Funcion para calcular makespan

def makespan(secuencia,tiempos):
    m = len(tiempos[0])  # Numero de maquinas
    n = len(secuencia)     # Numero de trabajos
    
    # Matriz organizada de acuerdo a la secuencia
    matriz = [[tiempos[i][j] for j in range(m)] for i in secuencia]
    
    make = [[0 for j in range(m)] for i in secuencia]
    
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

# =============================================
# HEURISTICA PROPUESTA
# =============================================

t_inicial = time.time()

# PRIMER PASO: ordenar de mayor a menor los trabajos de acuerdo a su tiempo total de procesamiento

# Calcular el tiempo total de procesamiento en todas las maquinas por cada trabajo.
T = []
contador = 0
for i in t:
    T.append((contador,sum(i)))
    contador +=1
    
# Organizar de mayor a menor y hallar secuencia de trabajos
T.sort(key = lambda x: x[1], reverse = True)

secuenciaT = []
for i in T:
    secuenciaT.append(i[0])

# SEGUNDO PASO: aplicar forward y backward shift mechanism y escoger mejor secuencia

def insercion(lista, ins_pos1, en_pos2):   
    x=lista.pop(ins_pos1)
    lista.insert(en_pos2,x)  
    return lista

# Realizar forward shift mechanism
Cmax = 999999
secuencia = secuenciaT[:]
mejorSecuencia = []
for i in range(1,len(secuenciaT)):
    lista = insercion(secuencia, 0,i)
    secuencia = secuenciaT[:]
    cMax = makespan(lista, t)
    if cMax < Cmax:
        mejorSecuencia = lista[:]
        Cmax = cMax

# Realizar backward shift mechanism
for j in range(1,len(secuenciaT)):
    lista = insercion(secuencia,j,0)
    secuencia = secuenciaT[:]
    cMax = makespan(lista, t)
    if cMax < Cmax:
        mejorSecuencia = lista[:]
        Cmax = cMax

# TERCER PASO: escoger la mejor secuencia para los primeros dos trabajos de la secuencia hasta ahora encontrada
k = 1
pareja = mejorSecuencia[0:2]
parejaInv = pareja[:]
parejaInv.reverse()

CmaxPareja = makespan(pareja,t)
CmaxParejaInv = makespan(parejaInv,t)

if CmaxPareja < CmaxParejaInv:
    semiSecuencia = pareja[:]
else:
    semiSecuencia = parejaInv[:]
    
# CUARTO PASO: seguir ubicando cada pareja 
ultimaSecuencia = []
numParejas = int(n/2)
for k in range(2,numParejas*2,2):
    parejaActual = mejorSecuencia[k:k+2]
    primero = parejaActual[0]
    segundo = parejaActual[1]
    mejorCmax = 99999
    # Ubica la pareja en la mejor ubicación de la secuencia actual
    for i in range(len(semiSecuencia)+1):
        semiSecuenciaActual = semiSecuencia[:]
        semiSecuenciaActual.insert(i,parejaActual)
        actual = flatter(semiSecuenciaActual)
        CmaxActual = makespan(actual,t)
        if CmaxActual < mejorCmax:
            ultimaSecuencia = actual[:]
            mejorCmax = CmaxActual
    semiSecuencia = ultimaSecuencia.copy()
    # Reubica el primer trabajo de la pareja k
    for j in range(len(semiSecuencia)+1):
        semiSecuencia.remove(primero)
        semiSecuencia.insert(j,primero)
        CmaxActual = makespan(semiSecuencia,t)
        if CmaxActual < mejorCmax:
            ultimaSecuencia = semiSecuencia[:]
            mejorCmax = CmaxActual
    semiSecuencia = ultimaSecuencia.copy()
    # Reubica al segundo trabajo de la pareja k
    for x in range(len(semiSecuencia)+1):
        semiSecuencia.remove(segundo)
        semiSecuencia.insert(j,segundo)
        CmaxActual = makespan(semiSecuencia,t)
        if CmaxActual < mejorCmax:
            ultimaSecuencia = semiSecuencia[:]
            mejorCmax = CmaxActual
    semiSecuencia = ultimaSecuencia.copy()

t_total = time.time()-t_inicial

# ==================================
# IMPRESION DE RESULTADOS
# ==================================
F0 = makespan(ultimaSecuencia,t)
solucion = []
for i in ultimaSecuencia:
    solucion.append(i+1)
print('\n')
print("La secuencia recomendada es: " + str(solucion))
print("Cmax: " + str(F0))
print(t_total/60)
        
        






















