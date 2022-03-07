# -*- coding: utf-8 -*-
"""
Created on Tue Jan 18 14:18:36 2022

@author: AARON.RAMIREZ

Este proceso es conveniente hacerlo una vez que el informante 
ha contestado el cuestionario para verificar que sus respuestas
coincidan con la validación establecida en proceso "encontrar"

Se necesita la guía y el censo contestado
"""
from utilidades.frames import Frame
import pandas as pd 
import openpyxl as op

documento = 'p_formula_blanco_vprel.xlsx'
guia = pd.read_csv('PR_p_formula.csv')

datos = guia



libro = op.open(documento,data_only=True)



del datos['formulas']
datos = datos.loc[datos['ID']!='W']

datos = datos.reset_index(drop=True)

cord = datos['coordenada']
secciones = datos['seccion'].unique()

def NS(comparador,referente):
    n = ['NS','ns']
    a = ['NA','na']

    if referente in a and comparador in a:
        return 0
    if referente in a and comparador not in a:
        return 1
    if referente not in a and comparador in a:
        return 0
    if referente in n and comparador in n:
        return 0
    if referente == 0 and comparador in n:
        return 1
    if referente in n and comparador >= 0:
        return 1
    if referente > 0 and comparador in n:
        return 0
    
    


d = Frame()
d.n_col('resultados',[])
d.n_col('valores',[])
for seccion in secciones:
    
    # try:
    si = libro[seccion]
    fila = 0
    for i in cord:
        
        if seccion == datos['seccion'][fila]:
            
            a = i.split(',')
           
            va = []
            val = []
            if len(a) > 1:
                suma = 0
                ns = []
                na = []
                nsd = ['NS','ns']
                nad = ['NA','na']
                lista = []
                for co in a:
                    valor = si[co]
                    if not valor.value:
                        valor = 0
                        lista.append(valor)
                    else:
                        lista.append(valor.value)
                for v in lista:
                    if v in nsd:
                        ns.append(1)
                        v = 0
                    if v in nad:
                        na.append(1)
                        v = 0
                    suma += v
                if suma == 0 and 1 in ns:
                    suma = 'NS'
                if suma == 0 and 1 in na:
                    suma = 'NA'
                va.append(str(suma))
                val.append(suma)
            else:
                for co in a:
                    valor = si[co]
                    if not valor.value:
                        va.append(str(0))
                        val.append(0)
                    else:
                        va.append(str(valor.value))
                        val.append(valor.value)
            
            d.n_col(str(datos['ID'][fila]),val)
            #el nombre d epregunta debe comenzar con un M1 o algo para hacer referencia al modulo al que pertenece
            an = ','.join(va)
            d.add(an,'valores')
            # datos['valor'][fila] = an queda mejor con la opcion de abajo por el mensaje de error
            # datos.iat[fila,datos.columns.get_loc('valor')] = an
            # print(d.buscar_col(str(datos['pregunta'][fila])))
            n = ['NS','ns','NA','na']
            if datos['operacion'][fila] == 'igual':
                
                fila1 = 0
                resultado = 'NA'
                compara = d.buscar_col(str(datos['ID'][fila]))
                for val in d.buscar_col(str(datos['comparacion'][fila])):
                    
                    if val in n or compara[fila1] in n:
                        resultado = NS(compara[fila1],val)
                    else:
                        if val == compara[fila1]:
                            resultado = 0
                        else:
                            resultado = 1
                    fila1 +=1
                d.add(resultado, 'resultados')
                
            if datos['operacion'][fila] == 'mayor':
                
                fila1 = 0
                resultado = 'NA'
                compara = d.buscar_col(str(datos['ID'][fila]))
                for val in d.buscar_col(str(datos['comparacion'][fila])):
                    
                    if val in n or compara[fila1] in n:
                        
                        resultado = NS(compara[fila1],val)
                    else:
                        if val <= compara[fila1]:
                            resultado = 0
                        else:
                            resultado = 1
                    fila1 +=1
                d.add(resultado, 'resultados')
            
            if datos['operacion'][fila] == 'menor':
                
                fila1 = 0
                resultado = 'NA'
                compara = d.buscar_col(str(datos['ID'][fila]))
                for val in d.buscar_col(str(datos['comparacion'][fila])):
                    
                    if val in n or compara[fila1] in n:
                        resultado = NS(compara[fila1],val)
                    else:
                        if val >= compara[fila1]:
                            resultado = 0
                        else:
                            resultado = 1
                    fila1 +=1
                d.add(resultado, 'resultados')
            
            if datos['operacion'][fila] not in ['menor','mayor','igual']:                    
                d.add('NA', 'resultados')
                
        else:
            
            pass
            # d.add('pendiente','valores')
            # d.add('pendiente', 'resultados')
        
        fila += 1
            
    # except:
    #     print('mala seccion')
    #     pass
    # print(len(cord),fila)
                  
        
    
datos['valor'] = d.buscar_col('valores')
datos['resultado'] = d.buscar_col('resultados')

doc = documento.split('.')
datos.to_csv(f'resultado{doc[0]}.csv',index=False)

# si = libro['Hoja1']

# a = si['B6']
# b = si['C6']

# if a.value == b.value:
#     print('Los valores coinciden')
# else:
#     print('Los valores no coinciden')
