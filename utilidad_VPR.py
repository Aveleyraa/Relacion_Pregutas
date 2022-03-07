# -*- coding: utf-8 -*-
"""
Created on Wed Feb 23 10:14:14 2022

@author: AARON.RAMIREZ
"""


def p_rel(guia, excel, pagina):
    """
    

    Parameters
    ----------
    guia : dataframe de pandas, con los elementos tal y 
    como los genera proceso llamado "encontrar.py"
    
    excel : la página del archivo excel a validar con openpyxl.
    
    pagina : nombre de la página del excel a validar.

    Returns
    -------
    None.

    """
    orden = categorizar(guia,pagina)
    
    for w in orden:
        escribir(w,excel)
    
    return

def escribir(lista,excel):
    fila = lista[0]
    columna = 32 #porque AF es 32
    for formula in lista[1:]:
        excel.cell(row=fila,column=columna,value=formula)
        columna += 1
    return


def categorizar(datos,pag):
    rl = []
    mdf = datos.loc[datos['seccion']==pag]
    w = mdf.loc[mdf['ID']=='W']
    rangoW = []
    for val in w['coordenada']:
        fila = ''
        for caracter in val:
            if caracter.isnumeric():
                fila += caracter
        rangoW.append(int(fila))
    rangoW.sort()
    mdf1 = mdf.loc[mdf['ID']!='W']
    mdf1 = mdf1.loc[mdf1['operacion']!='ref']
    numeros = []
    for val in mdf1['coordenada']:
        if ',' in val: #para omitir sumas
            pass
        else:
            fila = ''
            for caracter in val:
                if caracter.isnumeric():
                    fila += caracter
            numeros.append(int(fila))
        
    ldel = []
    for rango in rangoW:
        ap = []
        for numero in numeros: #conseguir numeros para los rangos
            if numero > rango:
                ap.append(numero)
        ap.sort()
        ldel.append(ap)
    for lista in ldel: #depurar numeros de los rangos
        c = 0
        for val in lista[1:]:
            resta = val - lista[c]
            if resta > 7:
                lista.remove(val)
    
    mdf1 = mdf1.reset_index(drop=True)
    
    medidor = 0
    for lista in ldel:
        
        sul = []
        sul.append(rangoW[medidor])
        comparar = list(set(lista))
        fila = 0
        for val in mdf1['coordenada']:
            
            for co in comparar:
                if val.endswith(str(co)):
                    sul.append(mdf1['formulas'][fila])
            fila += 1
        rl.append(sul)
        medidor += 1
                
    return rl

#primer paso: ordenar y agrupar coordenadas con referente W y excluir las que no
#Segundo paso: hacer la validación
#Tercer paso: quizá meter condicionales?

