# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 12:07:33 2022

Mis propios datamframes

@author: AARON.RAMIREZ
"""

class Frame():
    
    def __init__(self):
        self.core = {}

    def n_col(self, nombre, contenido):
        "nombre = string, contenido = lista"
        if type(contenido) == list and type(nombre) == str:
            self.core[nombre] = contenido
        else:
            print('Revisar argumentos de entrada para metodo n_col; \n contenido = lista, nombre = string')
        
    def add(self,valor, columna_nombre):
        "valor es lo que sea que se quiera agregar a fila en columna"
        if type(columna_nombre) == str:
            if columna_nombre in self.core:
                self.core[columna_nombre].append(valor)
            else:
                print('La columna no se encuentra en el objeto. Metodo add')
        else:
            print('El nombre de la columna debe ser un string')
    
    def buscar_col(self,columna_nombre):
        if type(columna_nombre) == str:
            if columna_nombre in self.core:
                return self.core[columna_nombre]
            else:
                print(f'{columna_nombre} no se encuentra en el objeto. Metodo buscar')
        else:
            print('El nombre de la columna debe ser un string')

alfa = Frame()

alfa.n_col('1.2', [1,5,8])
alfa.add('kilo','1.2')