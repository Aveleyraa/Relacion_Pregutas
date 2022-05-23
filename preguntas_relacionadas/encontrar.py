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


def start_revision(self):


    libro = filedialog.askopenfilename()
    base = {}

    #Aqui se inicia el proceso de lectura para cada hoja del documento
    pags = pd.ExcelFile(libro).sheet_names

    for pag in pags:
        pagina = pag
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

    #guardar

    nom_s = libro.split('.')
    original.to_csv(f'PR_{nom_s[0]}.csv',index=False)



