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
from openpyxl.styles import PatternFill
from encontrar import getnum

documento = 'CNIFJ_2022_M1_S3_Resp (1).xlsx'
guia = pd.read_csv('PR_01_3_CNIJF_2022_M1_S3_V3(03dic21)_Act21Ene (1).csv')

datos = guia


libro = op.open(documento,data_only=True)


#Se borra columna de formulas porque no se necesita aquí
del datos['formulas'] 

#se filtran las marcas que no sean de escritura
datos = datos.loc[datos['ID']!='W']

#resetar el index es importante para evitar errores al iterar más adelante
datos = datos.reset_index(drop=True)

cord = datos['coordenada']
secciones = datos['seccion'].unique()

def NS(comparador,referente):
    """
    

    Parameters
    ----------
    comparador : str
        el valor detectado en excel, y que tiene dentro ns o na.
    referente : str
        el valor detectado en excel, y que tiene dentro ns o na.

    Returns
    -------
    int
        regresa 1 si es error, 0 si todo en orden.
    
    Nota: No alterar el orden de los condicionales, ya que eso puede 
    generar errores

    """
    n = ['NS','ns']
    a = ['NA','na']
    
    if referente in a and comparador in a:
        return 0
    if referente in a and comparador not in a:
        return 1
    if referente not in a and comparador in a:
        return 1 #Error discutido con Paulina sobre NA´s 
    if referente in n and comparador in n:
        return 0
    if referente == 0 and comparador in n:
        return 1
    if referente in n and comparador >= 0:
        return 1
    if referente > 0 and comparador in n:
        return 0
    
def comparacion(ref,com,op):
    """
    

    Parameters
    ----------
    ref : int or float
        numero del referente
    com : int or float
        numero del comparador
    op : int
        numero de operacion; 1 menor o igual,
        2 mayor o igual,
        3 igual

    Returns int
    0 si cumple operacion
    1 si no cumple operacion
    -------
    None.

    """
    if op == 1:
        if com <= ref:
            return 0
        else:
            return 1
    
    if op == 2:
        if com >= ref:
            return 0
        else:
            return 1
    
    if op == 3:
        if com == ref:
            return 0
        else:
            return 1
        


d = Frame()
d.n_col('resultados',[])
d.n_col('valores',[])
for seccion in secciones: #Para revisar las secciones que estan en el archivo guia
    

    si = libro[seccion]
    fila = 0
    for i in cord: #se itera por cada elemento que hay en la base de datos desde su coordenada
        
        if seccion == datos['seccion'][fila]: # si la sección es la misma con la que se está trabajando se hace esto
            
            a = i.split(',') #Para verificar si es más de una cordenada o solo una
           
            va = []
            val = []
            if len(a) > 1: #para casos donde hay mas de una coordenada lleva un proceso especial para sumar los valores de todas las coordenadas
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
            else: #esto se hace en caso de que solo sea una coordenada
                for co in a:
                    valor = si[co]
                    if not valor.value:
                        va.append(str(0))
                        val.append(0)
                    else:
                        va.append(str(valor.value))
                        val.append(valor.value)
            
            d.n_col(str(datos['ID'][fila]),val)
            
            an = ','.join(va)
            d.add(an,'valores')
            n = ['NS','ns','NA','na']
            
            #comenzar a evaluar
            fila1 = 0
            resultado = 'NA'
            compara = d.buscar_col(str(datos['ID'][fila]))
            for val in d.buscar_col(str(datos['comparacion'][fila])):
                try:
                    val = float(val)
                except:
                    pass
                try:
                    compara[fila1] = float(compara[fila1])
                except:
                    pass
                if val in n or compara[fila1] in n:
                        resultado = NS(compara[fila1],val)
                else:
                #Cuatro condicionales para evaluar de acuerdo al tipo de comparación
                    if datos['operacion'][fila] == 'igual':
                        resultado = comparacion(val,compara[fila1],3)
                    if datos['operacion'][fila] == 'mayor':
                        resultado = comparacion(val,compara[fila1],2)
                    if datos['operacion'][fila] == 'menor':
                        resultado = comparacion(val,compara[fila1],1)
                    if datos['operacion'][fila] not in ['menor','mayor','igual']:
                        resultado = 'NA'
                fila1 +=1
            d.add(resultado, 'resultados')
                
        else:
            
            pass
            
        
        fila += 1
            
   
                  
        
#Integrar resultados de comparaciones, junto con los valores detectados en el excel ya contestado al dataframe
datos['valor'] = d.buscar_col('valores')
datos['resultado'] = d.buscar_col('resultados')


#Marcar celdas con errores
#esto no es tan buena idea debido a que el archivo se abre con solo valores, por lo que al guardarlo se pierden las formulas
# cafe =PatternFill('solid',start_color='D35400',end_color='D35400')
# c = 0
# for error in datos['resultado']:
#     if error == 1:
#         cor = datos['coordenada'][c]
#         cr = cor.split(',')
#         seccion = datos['seccion'][c]
#         hoja = libro[seccion]
#         for corde in cr:
#             a1 = hoja[corde]
#             a1.fill = cafe
#     c += 1
# libro.save(documento)

#hacer columna de preguntas para indicar el numero de pregunta donde hay error
datos['preguntas'] = 'NA'


def preguntas(hopan):
    "regresa una lista con la coordenada del inicio de preguntas. Es usada en funcion espacio"
    c=0
    ind=[]
    for i in hopan['Unnamed: 0']:
        a = pd.isna(hopan['Unnamed: 0'][c])
        if a == False:
            ind.append(c)
        c+=1
    return ind



for seccion in secciones:
    df = pd.read_excel(documento,sheet_name=seccion)
    pre = preguntas(df)
    fila = 0
    for r in datos['resultado']:
        if r == 1:
            if datos['seccion'][fila] == seccion:
                cor = datos['coordenada'][fila]
                cor = cor.split(',')
                fi = []
                for co in cor:
                    a = getnum(co)
                    fi.append(a)
                fi.sort()
                fp = 0
                kl = -1
                for p in pre:
                    if fi[0] < p:
                        fp = pre[kl]
                        break
                    kl += 1
                hoja = libro[seccion]
                fp += 2
                ncor = f'A{fp}'
                npreg = hoja[ncor].value
                datos['preguntas'][fila] = npreg
        fila += 1
                
    

#guardar un nuevo dataframe con las columnas de valor y resultado
doc = documento.split('.')
datos.to_csv(f'resultado{doc[0]}.csv',index=False)

