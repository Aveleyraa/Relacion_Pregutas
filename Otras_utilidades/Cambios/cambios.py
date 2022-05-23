# -*- coding: utf-8 -*-
"""
Created on Fri Feb  4 12:02:07 2022

@author: AARON.RAMIREZ
"""

import pandas as pd
import ntpath

def cambios(libro, libro2):
    
    pags = pd.ExcelFile(libro).sheet_names
        
        #Antes de correr el ciclo, asegurarse de que pags tenga todas las hojas
        #en las que se van a escribir validaciones con este método.
        

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



def path_leaf(path):
    #funcion para detectar el path del archivo seleccionado
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

