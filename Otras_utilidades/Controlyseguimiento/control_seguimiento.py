# -*- coding: utf-8 -*-
"""
Created on Thu Mar 31 16:26:00 2022

@author: AARON.RAMIREZ
"""

import pandas as pd
import openpyxl as op

def con_y_seg(book, observaciones, nrev, fepro, fere, lev, pm):
    
    hoja = book['M1']

    hojap = 'M1_S1' #nombre de la hoja de excel donde están las observaciones

    datos = pd.read_excel(observaciones,sheet_name=hojap)
    datos = datos.fillna(0)

    #nrev = input('Qué número de revisión es? ')
    #fepro = input('Cuál es la fecha programada de entrega? ')
    #fere = input('Cuál fue la fecha real de entrega? ')
    #lev = input('Cuál es el año del levantamiento? ')
    #pm = input('Cuántas preguntas tiene el módulo? ')
    dhabiles = '=NETWORKDAYS(H11,M11)'
    programa = datos['Unnamed: 3'][9]
    preguntas = []
    totobs = 0 #total de observaciones detectadas


    incons = [0,0,0,0,0,0,0,0,0] 

    fila = 0
    for obs in datos['Unnamed: 2'][16:]:
        if obs != 0:
            obs = str(obs)
            obs = obs.replace(' ','')
            contenido = obs.split(',')
            preguntas += contenido
            totobs += 1
            tipo = datos['Unnamed: 3'][16+fila]
            if 'responder' in tipo:
                incons[0] += len(contenido)
            if 'incompletas' in tipo:
                incons[1] += len(contenido)
            if 'especifique' in tipo:
                incons[2] += len(contenido)
            if 'aritmética' in tipo:
                incons[3] += len(contenido)
            if 'consistencia' in tipo:
                incons[4] += len(contenido)
            if 'registro' in tipo:
                incons[5] += len(contenido)
            if 'justificación' in tipo:
                incons[6] += len(contenido)
            if 'Diferencias' in tipo:
                incons[7] += len(contenido)
            if 'explicativo' in tipo:
                incons[2] += len(contenido)
            if 'Otra' in tipo:
                print('otra insonsistencia detectada')

            fila += 1

    uni = list(set(preguntas)) 
    uni.sort()      
    totp = len(uni) #total de preguntas con observaciones

    if preguntas:
        osi = 'Si'
    if not preguntas:
        osi = 'No'
    #imprimir en el formato los datos

    hoja['C9'] = int(lev)
    hoja['G9'] = programa
    hoja['C11'] = nrev
    hoja['H11'] = fepro
    hoja['M11'] = fere
    hoja['G13'] = dhabiles
    hoja['K13'] = datos['Unnamed: 3'][12]
    hoja['C17'] = osi
    if osi == 'Si':
        hoja['C19'] = totobs
    hoja['C21'] = '=K33'
    hoja['J27'] = pm
    hoja['J28'] = int(pm) - int(incons[0])
    hoja['J29'] = incons[0]
    hoja['J30'] = '=C19'
    hoja['J31'] = totp

    c = 0
    for inco in incons:
        fi = 27 + c
        hoja[f'G{fi}'] = inco
        c += 1

    return book