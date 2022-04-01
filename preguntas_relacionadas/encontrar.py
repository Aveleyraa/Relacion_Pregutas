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


libro = '01_DIMJA_24ene2022_marcas (1).xlsx'


base = {}




def genp():
    """
    regresa res que es una lista con letras mayusculas y minusculas 
    para buscar coincidencias en el documento excel (las marcas)
    ejemplo ["A1","b1"...]
    """
    abc1 = list(string.ascii_uppercase)
    abc = list(string.ascii_lowercase)
    nn = ['Ñ','ñ'] 
    l = abc1 + abc + nn
    res = []
    for i in l:
        for d in range(1,10):
            res.append(i+str(d))
        
    return res

posibles = genp() + ['W'] #variable con la lista para buscar letras con numeros en el documento

def filtro(cadena):
    """
    

    Parameters
    ----------
    cadena : str 

    Returns
    -------
    str
        Regresa una string dependiendo el signo que encuentre en la cadena
        que se pasa como argumento de entrada. En caso de que no detecte
        ninguno, regresa ref, que ace alución a las marcas que son referentes
        y no necesitan ser comparados

    """
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
    """
    

    Parameters
    ----------
    sec : str
        Es el nombre de la hoja de excel que se va a leer.
    datos : dataframe pandas
        es la hoja de excel a leer, expresada en un dataframe de pandas.

    Returns
    -------
    entot : dic
        regresa un diccionario que contiene filas, columnas y sección
        donde se encontraron las letras de variable "posibles". además 
        contiene dentro de sí, otros diccionarios referentes alas letras
        que encontró, determianndo si son referentes o comparadores,
        la id a la que hacen referencia los comparadores, así como su
        operacion de acuerdo al signo detectado dentro de la string.

    """
    mdf = datos
    reg = 0
    cont = 0
    entot = {'fila':[],'columna':[],
             'sec':[]}
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
    """
    

    Parameters
    ----------
    sec : str
        nombre de la hoja de excel a leer.
    datos : dataframe de pandas
        la hoja de excel a leer expresada en dataframe de pandas.

    Returns
    -------
    d_salida : dic
        Regresa diccionario con las coordenadas expresadas a manera de 
        documento en excel, ejemplo "A25" para referir a una celda que
        está en columna A, fila 25. También tiene la sección donde está
        sacando esas coordenadas y la id de las macas que encontró. 
        La id en este caso es la marca como se ha puesto en el documento,
        por ejemplo "a1.1>"

    """
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
    "esta funcion ordena el diccionario que da de salida la funcion coordenadas"
    b = {'seccion':[],'coordenada':[],
         'comparacion':[],'operacion':[],'ID':[]}
    for k in di:
        if 'sec' not in k and 'Op' not in k and 'id' not in k: #se itera solo para el elemento en el diccionario que corresponde a la letra buscada en la marca, no a su id, ni secion ni operacion
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
    """
    

    Parameters
    ----------
    cadena : str
        operacion que se va a realizar para comparar.

    Returns
    -------
    str
        elementos para realizar formula que se escribirá en excel. 

    """
    if cadena == 'menor':
        return '<='
    if cadena == 'mayor':
        return '>='
    if cadena == 'igual':
        return '='
    else:
        return 'posible mala referencia'

def clasif(cadena):
    """
    

    Parameters
    ----------
    cadena : str
        operacion que se va a realizar para comparar.

    Returns
    -------
    str
       regresarlo en palabras es importante para el proceso de revision.

    """
    if '<' in cadena:
        return 'menor'
    if '>' in cadena:
        return 'mayor'
    if '=' in cadena:
        return 'igual'
    else:
        return 'ref'

def formulaS(cadena,seccion):
    """
    

    Parameters
    ----------
    cadena : str
        cadena con más de una coordenada.
    seccion : str
        nombre de la hoja de excel.

    Returns
    -------
    formula : str
        Dado que hay más de una coordenada, se tiene que hacer un proceso
        adicional para generar una formula que nos dé el valor concreto
        de la suma de los valores que se detecten en las coordenadas
        que están siendo pasadas como argumento de esta funcion. Regresa
        por lo tanto una formula que hace eso.

    """
    if seccion != '':
        ad = seccion+'!'
    else:
        ad = seccion
    c = cadena.split(',')
    r = f'COUNTIF({ad}{c[0]}:{c[-1]},"NS")'
    ca = f'{ad}{c[0]}:{c[-1]}'
    bl = f'ISBLANK({ad}{c[0]})'
    o = f'COUNTIF({ad}{c[0]}:{c[-1]},"NA")'
    for co in c[1:]: 
        bl += f',ISBLANK({ad}{co})' #Esto porque es blanco solo funciona con una celda
    #metodo para coordenadas si no fueran continuos
    # r = f'COUNTIF({ad}{c[0]},"NS")'
    # ca = f'{ad}{c[0]}'
    # bl = f'ISBLANK({ad}{c[0]})'
    # o = f'COUNTIF({ad}{c[0]},"NA")'
    # for co in c[1:]: 
    #     r += f'+COUNTIF({ad}{co},"NS")'
    #     ca += f',{ad}{co}'
    #     bl += f',ISBLANK({ad}{co})'
    #     o += f'+COUNTIF({ad}{co},"NA")'
      
    formula = f'IF(AND(SUM({ca})=0,{r}>0),"NS",IF(AND(SUM({ca})=0,{o}>0),"NA",IF(AND({bl}),"",SUM({ca}))))'
    return formula

def getnum(cad):
    "regresa el numero de fila de una coordenada de excel"
    numero = ''
    for caracter in cad:
        if caracter.isnumeric():
            numero += caracter
    numero = int(numero)
    return numero

def sumco(co,num):
    """
    

    Parameters
    ----------
    co : str
        coordenada de excel ejemplo A25.
    num : int
        numero que se va a sumar a la fila de la coordenada.

    Returns
    -------
    cor : str
        nueva coordenda con el num sumado a la fila de la coordenada de 
        entrada.

    """
    letra = ''
    fila = getnum(co)
    for caracter in co:
        if caracter.isalpha():
            letra += caracter
    cor = f'{letra}{fila+num}'
    return cor

def columnas(unicos,base,secc):
    """
    

    Parameters
    ----------
    unicos : list
        Lista de valores unicos en donde se detectó existencia 
        del caracter ":"
    base : dic
        Diccionario donde están conenidos todos los elementos registrados
        de una hoja de excel.
    secc : str
        nombre de la hoja de excel.

    Returns
    -------
    base : dic
        Regresa el diccionario de entrada pero modificado ya que agrega 
        las celdas contenidas en las columnas marcadas. Las marcas con ":"
        representan una columna a comparar con otra, donde en realidad cada
        fila de esa columna tiene que ser comparada con la fila de otra.
        Por esa razón se generan las referencias necesarias a cada fila dentro
        de las columnas que fueron marcadas. Además, se hace el borrado del
        caracter ":" para no generar errores en los procesos siguentes
        de creación de formulas.

    """
    for columna in unicos:
        
        seccion = secc
        coord = []
        ide = ''.join(l for l in columna if l != ':')
        op = clasif(columna)
        c = 0
        indices = []
        for ele in base['ID']:
            
            if ele == columna:
                coord.append(base['coordenada'][c])
                indices.append(c)
            c += 1
        cor = coord[0]
        a1 = getnum(coord[0])
        a2 = getnum(coord[1])
        resta = a2-a1
        ide1 = ide
        if '.' in ide:
            iterar = ide.split('.')
            n = iterar[0]
            ide1 = n
            
        d = {'seccion':seccion,'coordenada':cor,
             'comparacion':ide1,'operacion':op,'ID':ide}
        integrar = [d]
        for i in range(1,resta+1):
            e = {'seccion':seccion,'coordenada':cor,
                 'comparacion':ide,'operacion':op,'ID':ide}
            e['coordenada'] = sumco(coord[0],i)
            e['ID'] = ide+str(i)
            if '.' in e['comparacion']:
                e['ID'] = ide
                iterar = e['comparacion'].split('.')
                n = iterar[0]
                e['comparacion'] = n + str(i)
            else:
                e['comparacion'] = ide+str(i)
            integrar.append(e)
        for ind in reversed(indices):
            for lla in base:
                base[lla].pop(ind)
        for fila in reversed(integrar):

            for ke in fila:
                base[ke].insert(indices[0],fila[ke])
     
    return base


#Aqui se inicia el proceso de lectura para cada hoja del documento


pags = pd.ExcelFile(libro).sheet_names

saltar = [
    'Índice',
    'Presentación',
    'Informantes',
    'Participantes',
    'Glosario']
for pag in pags:
    pagina = pag
    if pagina not in saltar:
        data = pd.read_excel(libro,sheet_name=pagina,engine='openpyxl')
        a = cordenadas(pag,data)
        r = nframe(a)
        base[pag] = r

#sacar sumas y columnas a los diccionarios generados para cada hoja donde se encontró una marca

for k in base:
    
    sua = []
    colum = []
    for ele in base[k]['ID']:
        if '+' in ele:
            sua.append(ele)
        if ':' in ele:
            colum.append(ele)
    sua1 = list(set(sua))
    
    if colum:
        column = list(set(colum))
        base[k] = columnas(column,base[k],k)
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
    
   

#Hacer el dataframe


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
                if sec != sec1: 
                    b = sec1+'!'+b
                if ',' in b:
                    if sec == sec1:
                        b = formulaS(b,'')
                    else: #referente de formula a otra hoja
                        b = formulaS(b,sec1)
                    
            else:
                pass
            filac += 1
        
        if a == 'posible mala referencia':
            signos = ['<','>','=']
            rt = 0
            for signo in signos:
                if signo in original['ID'][fila]: #esta comprobación es para las validaciones de columnas donde se usa el ":"
                    rt = 1
            if original['operacion'][fila] == 'ref' and rt == 1:
                formulas.append('NA')
            else:
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
    if len(original['comparacion'][fila])>2:
        original['formulas'][fila] = 'NA'
    fila += 1
borrar = original[original['formulas']=='Borrar'].index
original = original.drop(borrar)

# poner referentes al inicio de cada seccion 

secciones = original['seccion'].unique()

contenedor = pd.DataFrame()

for seccion in secciones:
    nframe = original.loc[original['seccion']==seccion]
    nframe = nframe.sort_values(by=['operacion'],ascending=False)
    contenedor = pd.concat([contenedor,nframe])
contenedor = contenedor.reset_index(drop=True)

# encontrar referentes duplicados
dup = contenedor['ID'].value_counts()

fila = 0

for i in contenedor['ID']:
    if contenedor['operacion'][fila] == 'ref':
        if dup[i] > 1:
            contenedor['formulas'][fila] = 'repetido'
    fila += 1

#guardar

nom_s = libro.split('.')
contenedor.to_csv(f'PR_{nom_s[0]}.csv',index=False,encoding='latin1')



