# -*- coding: utf-8 -*-
"""
Created on Wed Feb 23 09:58:20 2022

@author: AARON.RAMIREZ

Este proceso debe ser posterior a "encontrar" pero previo a 
si quiera enviar el cuestionario validado, ya que aquí es donde se 
imprimen las fórmulas de validación en el cuestionario


SE necesita guía y censo para validar (puede ser un 
   cuestionario ya con validaciones o en blanco, según se disponga)

es importante que existan marcas "W" en el cuestionario para que se 
pueda hacer la escritura de forma correcta
"""

import pandas as pd
import openpyxl as op
from utilidades.utilidad_VPR import p_rel

guia = pd.read_csv('PR_02_CNDHF_2022_M2_marcas.csv')

archivo = '02_CNDHF_2022_M2_Capacitación difusión defensa y protección de los derechos humanos_VF(27Abr22)_validaciones.xlsx'#Tiene que ser cuestionario para meter validaciones



book = op.load_workbook(archivo)

pags = book.sheetnames


#iterar por cada una de las hojas de excel para escribir las validaciones

for pa in pags:
    
    pagina = pa
    
    shi = book[pagina]
    if pa in guia.seccion.values:
        
        p_rel(guia, shi, pa)
    else:
        pass

nom_s = archivo.split('.')
book.save(f'{nom_s[0]}_vprel.xlsx')