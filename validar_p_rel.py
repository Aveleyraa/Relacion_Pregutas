# -*- coding: utf-8 -*-
"""
Created on Wed Feb 23 09:58:20 2022

@author: AARON.RAMIREZ

Este proceso debe ser posterior a "encontrar" pero previo a 
si quiera enviar el cuestionario validado, ya que aquí es donde se 
imprimen las fórmulas de validación en el cuestionario


SE necesita guía y censo para validar (puede ser un 
   cuestionario ya con validaciones o en blanco, según se disponga)
"""

import pandas as pd
import openpyxl as op
from utilidad_VPR import p_rel

guia = pd.read_csv('PR_p_formula.csv')

archivo = 'p_formula_blanco.xlsx'#Tiene que ser cuestionario para meter validaciones



book = op.load_workbook(archivo)

pags = book.sheetnames


# para_pags = ['Secc']
# pags = []
# for val in para_pags:
#     pags1 = [pag for pag in p if val in pag]
#     pags += pags1
  

# pags = ['Hoja1','Hoja2']
# pags = ['CNIJF_2022_M1_Secc1_Sub5']
for pa in pags:
    
    pagina = pa
    
    shi = book[pagina]
    if pa in guia.seccion.values:
        
        p_rel(guia, shi, pa)
    else:
        pass

nom_s = archivo.split('.')
book.save(f'{nom_s[0]}_vprel.xlsx')