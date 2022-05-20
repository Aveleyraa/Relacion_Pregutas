# -*- coding: utf-8 -*-
"""
Created on Fri Feb  4 12:02:07 2022

@author: AARON.RAMIREZ
"""

import pandas as pd
import openpyxl as op

# archivo = '01_CNSIPEF_2022_M1_Estructura organizacional y recursos_VF(21Sep21)_Act05Nov'

libro = 'CNIJF_2022_M1_S3_VF_R2.1.xlsx'
libro2 = 'CNIJF_2022_M1_S3_rev.xlsx'

# book = op.load_workbook(libro)

# p = book.sheetnames


# para_pags = ['Anexo']
# pags = []
# for val in para_pags:
#     pags1 = [pag for pag in p if val in pag]
#     pags += pags1

# pags = ['Anexo '+str(i) for i in range(1,14)]

pags = pd.ExcelFile(libro).sheet_names
  
"""
Antes de correr el ciclo, asegurarse de que pags tenga todas las hojas
en las que se van a escribir validaciones con este método.
"""

for pa in pags:
    
    pagina = pa
    
    
    shet = pd.read_excel(libro,sheet_name=pagina,engine='openpyxl')
    shet =shet.fillna(0)
    shet = shet.iloc[:,0:31]
    compar = pd.read_excel(libro2,sheet_name=pagina,engine='openpyxl')
    compar = compar.fillna(0)
    compar = compar.iloc[:len(shet),0:31]
    try:
        if shet.equals(compar):
            print(f'{pa} es igual en ambos documentos')
        else:
            
            res = shet==compar
            res.to_csv(f'{pa}.csv',index=False)
            print(f'{pa} tiene diferencias pero puede compararse. Revisar archivo generado para ver cambios')
    except:
        print(f'Sección {pa} no puede ser comparada porque tiene dimensiones diferentes')



