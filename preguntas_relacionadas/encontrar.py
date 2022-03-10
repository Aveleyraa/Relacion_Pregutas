# -*- coding: utf-8 -*-
"""
Created on Fri Feb 18 13:26:03 2022

@author: AARON.RAMIREZ

Este es el primer paso de preguntas relacionadas.
Se trata de localizar las coordenadas y generar 
sus fórmulas para validacion

Se necesita cuestionario marcado nada más
"""

import pandas as pd
import string

libro = 'p_formula.xlsx'


base = {}




def genp():
    abc1 = list(string.ascii_uppercase)
    abc = list(string.ascii_lowercase)
    l = abc1[:12] + abc[:12]
    res = []
    for i in l:
        for d in range(1,10):
            res.append(i+str(d))
        
    return res

posibles = genp() + ['W']

def filtro(cadena):
    a = '>'
    b = '<'
    c = '='
    if a in cadena:
        return 'mayor'
    
    if b in cadena:
        return 'menor'
    
    if c in cadena:
        return 'igual'
    
    else:
        return 'ref'
    
    
def imagen(sec,datos):

    mdf = datos
    reg = 0
    cont = 0
    entot = {'fila':[],'columna':[],
             'comp':[],'sec':[],'id':[]}
    asig = {}
    asig1 = {}
    asid = {}
    for re in posibles:
        asig[re] = []
        asid[re+'id'] = []
        asig1[re+'Op'] = []
    for lista in mdf:
        vo = 0
        for col in datos[lista]:
            
            for l in posibles:
                
                try:
                    t = col.split('/')
                    for i in t:
                        
                        if l == i:
                            asig[l] = [reg,*asig[l]]
                            asid[l+'id'] = [i,*asid[l+'id']]
                            asig1[l+'Op'] = [filtro(i),*asig1[l+'Op']]
                        if i.startswith(l) and i!=l:
                            asig[l].append(reg)
                            asid[l+'id'].append(i)
                            asig1[l+'Op'].append(filtro(i))
                        if l in i:
                            entot['sec'].append(sec)
                            entot['fila'].append(vo)
                            entot['columna'].append(cont)
                            reg += 1
                    
                except:
                    pass
            vo+=1
        cont+=1
    borrar = [val for val in asig if not asig[val]]
    for i in borrar:
        del asig[i],asig1[i+'Op'],asid[i+'id']
    
    entot['asig'] = asig
    entot['asig1'] = asig1
    entot['asid'] = asid
    
    return entot

def cordenadas(sec,datos):
    abc1 = list(string.ascii_uppercase)+['AA','AB','AC','AD','AE']
    d = imagen(sec,datos)
    d_salida = {}
    sal = []
    c = 0
    for n in d['fila']:
        try:
            cor = abc1[d['columna'][c]]+str(n+2)
            sal.append(cor)
        except:
            sal.append('Fuera de margen')
        c +=1
    for ele in d['asig']:
        d_salida[ele] = []
        d_salida[ele+'sec'] = []
        # d_salida[ele+'id'] = []
        for v in d['asig'][ele]:
           d_salida[ele].append(sal[v]) 
           d_salida[ele+'sec'].append(d['sec'][v]) 
           # d_salida[ele+'id'].append(d['asid'][v])
    d_salida.update(d['asig1'])
    d_salida.update(d['asid'])
    
    return d_salida

def nframe(di):
    b = {'seccion':[],'coordenada':[],
         'comparacion':[],'operacion':[],'ID':[]}
    for k in di:
        if 'sec' not in k and 'Op' not in k and 'id' not in k:
            c = 0
            for val in di[k]:
                
                b['coordenada'].append(val)
                b['comparacion'].append(k)
                b['operacion'].append(di[k+'Op'][c])
                b['seccion'].append(di[k+'sec'][c])
                b['ID'].append(di[k+'id'][c])
                c += 1
    return b

def determinar(cadena):
    if cadena == 'menor':
        return '<='
    if cadena == 'mayor':
        return '>='
    if cadena == 'igual':
        return '='
    else:
        return 'posible mala referencia'

def clasif(cadena):
    if '<' in cadena:
        return 'menor'
    if '>' in cadena:
        return 'mayor'
    if '=' in cadena:
        return 'igual'
    else:
        return 'ref'

def formulaS(cadena,seccion):
    if seccion != '':
        ad = seccion+'!'
    else:
        ad = seccion
    c = cadena.split(',')
    r = f'COUNTIF({ad}{c[0]},"NS")'
    ca = f'{ad}{c[0]}'
    bl = f'ISBLANK({ad}{c[0]})'
    o = f'COUNTIF({ad}{c[0]},"NA")'
    for co in c[1:]:
        r += f'+COUNTIF({ad}{co},"NS")'
        ca += f',{ad}{co}'
        bl += f',ISBLANK({ad}{co})'
        o += f'+COUNTIF({ad}{co},"NA")'
        
    formula = f'IF(AND(SUM({ca})=0,{r}>0),"NS",IF(AND(SUM({ca})=0,{o}>0),"NA",IF(AND({bl}),"",SUM({ca}))))'
    return formula

"""
Documento

"""



# se = ['Hoja1','Hoja2']
pags = pd.ExcelFile(libro).sheet_names


for pag in pags:
    pagina = pag
    data = pd.read_excel(libro,sheet_name=pagina,engine='openpyxl')
    a = cordenadas(pag,data)
    r = nframe(a)
    base[pag] = r

#sacar sumas

for k in base:
    
    sua = []
    for ele in base[k]['ID']:
        if '+' in ele:
            sua.append(ele)
    sua1 = list(set(sua))
    
    for suma in sua1:
        
        seccion = k
        coord = []
        ide = ''.join(l for l in suma if l != '+')
        op = clasif(suma)
        c = 0
        for ele in base[k]['ID']:
            
            if ele == suma:
                coord.append(base[k]['coordenada'][c])
            c += 1
        cor = coord[0]
        for i in coord[1:]:
            cor += ','+i
        d = {'seccion':seccion,'coordenada':cor,
             'comparacion':ide,'operacion':op,'ID':ide}
        if '.' in d['comparacion']:
            iterar = d['comparacion'].split('.')
            n = iterar[0]
            d['comparacion'] = n
            for ke in d:
                base[k][ke].append(d[ke])
        else:
            for ke in d:
                base[k][ke].insert(0,d[ke])
    
        

#Hacer el frame


c = 0
for k in base:
    if c == 0:
        original = pd.DataFrame(base[k])
    else:
        ad = pd.DataFrame(base[k])
        original = pd.concat([original,ad],ignore_index=True)
    c += 1
    
#Hacer formulas   
formulas = []
fila = 0
for element in original['comparacion']:
    if element == original['ID'][fila]:
        formulas.append('NA')
    
    else:
        c = original['coordenada'][fila]
        if ',' in c:
            c = formulaS(c,'')
        a = determinar(original['operacion'][fila])
        sec = original['seccion'][fila]
        filac = 0
        for ele in original['ID']:
            if ele == element:
                b = original['coordenada'][filac]
                sec1 = original['seccion'][filac]
                if ',' in b:
                    
                    b = formulaS(b,'')
                if sec != sec1: #referente de formula a otra hoja
                    b = formulaS(b,sec1)
            else:
                pass
            filac += 1
        
        if a == 'posible mala referencia':
            formulas.append(a)
        else:
            try:
                form = f'=IF(AND({c}{a}{b},OR(AND(ISNUMBER({b}),ISNUMBER({c})),AND(ISBLANK({b}),ISBLANK({c})),OR(AND(ISBLANK({b}),{c}=""),AND(ISBLANK({c}),{b}="")))),0,IF(OR(AND({c}="NS",{b}>0,ISNUMBER({b})),AND({c}="NS",{b}="NS"),OR(AND({b}="NA",{c}="NA"),AND({b}="NA",ISBLANK({c})))),0,1))'
                formulas.append(form)
            except:
                form = 'No existe su referente'
                formulas.append(form)
    fila += 1

original['formulas'] = formulas

#depurar formulas y dataframe

fila = 0
for element in original['ID']:
    if '+' in element:
        if '.' in element:
            compa = ['>','<','=']
            cumple = 'no'
            for sig in compa:
                if sig in element:
                    cumple = 'si'
            if cumple == 'no':
                formulas.append('Mala referencia')
                original['formulas'][fila] = 'Mala referencia'
            if cumple == 'si':
                original['formulas'][fila] = 'Borrar'
        else:
            original['formulas'][fila] = 'Borrar'
    fila += 1
borrar = original[original['formulas']=='Borrar'].index
original = original.drop(borrar)

#guardar

nom_s = libro.split('.')
original.to_csv(f'PR_{nom_s[0]}.csv',index=False)



