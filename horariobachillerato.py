import pandas as pd  # type: ignore
import numpy as np  # type: ignore
from corregir_nombres import corregir_nombre

#leer base de datos
base = pd.read_excel('C:/Users/Admin/Desktop/GK/BASE DE DATOS 2025.xlsx', sheet_name = 'GK2025')
base['GRADO'] = base['GRADO'].astype(str)
base['ESTUDIANTE'] = base['ESTUDIANTE'].apply(corregir_nombre)

#leer listado general
bach = ['6','7','8','9','10','11']
listado = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name = 'g')
listado['GRADO'] = listado['GRADO'].astype(str)
listado.loc[listado['ESTUDIANTE'] == 'YARURO FONSECA JUAN ANTONIO', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'ACOSTA CASTAÑEDA JUAN CARLOS', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'ELVIRA DUARTE SARAH', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'MEDINA USECHE ANTONIO JOSÉ', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'OLAYA GONZÁLEZ CHRISTOPHER', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'PALMA ESTRADA JOSE MIGUEL', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'RÍOS LUGO JUAN DIEGO', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'VILLAMÍL GUANCHÉZ AMANDA ISABEL', 'GRADO'] = '6'
listado.loc[listado['ESTUDIANTE'] == 'YANDÚM BAUTISTA ANDRÉS FELIPE', 'GRADO'] = '6'
listado = listado[listado.iloc[:,2].isin(bach)]
listado.loc[listado['ESTUDIANTE'] == 'YARURO FONSECA JUAN ANTONIO', 'GRADO'] = '5'
listado.loc[listado['ESTUDIANTE'] == 'ACOSTA CASTAÑEDA JUAN CARLOS', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'ELVIRA DUARTE SARAH', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'MEDINA USECHE ANTONIO JOSÉ', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'OLAYA GONZÁLEZ CHRISTOPHER', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'PALMA ESTRADA JOSE MIGUEL', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'RÍOS LUGO JUAN DIEGO', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'VILLAMÍL GUANCHÉZ AMANDA ISABEL', 'GRADO'] = '4'
listado.loc[listado['ESTUDIANTE'] == 'YANDÚM BAUTISTA ANDRÉS FELIPE', 'GRADO'] = '4'

listado['ESTUDIANTE'] = listado['ESTUDIANTE'].apply(corregir_nombre)

#Leer base planeacion
base_planeacion = pd.read_excel('C:/Users/Admin/Desktop/GK/Horario/base planeacion.xlsx', sheet_name = 'bachillerato')
base_planeacion['Estudiante'] = base_planeacion['Estudiante'].apply(corregir_nombre)

a = input('INGRESE EL GRUPO DE INGLES :')

if a == 'F':
    orden = ['F','G','H','I','J','K','L','M']
if a == 'G':
    orden = ['G','H','I','J','K','L','M','F']
if a == 'H':
    orden = ['H','I','J','K','L','M','F','G']
if a == 'I':
    orden = ['I','J','K','L','M','F','G','H']
if a == 'J':
    orden = ['J','K','L','M','F','G','H','I']
if a == 'K':
    orden = ['K','L','M','F','G','H','I','J']
if a == 'L':
    orden = ['L','M','F','G','H','I','J','K']
if a == 'M':
    orden = ['M','F','G','H','I','J','K','L']

base_planeacion['INGLÉS'] = pd.Categorical(base_planeacion['INGLÉS'], categories=orden, ordered=True)
base_planeacion = base_planeacion.sort_values('INGLÉS')
base_planeacion = base_planeacion.reset_index(drop=True)


asignaturas_6_7= ['Biología','Química','Medio ambiente','Física',
                  'Historia', 'Geografía', 'Participación política','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Aritmética','Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

ciencias_6_7  = ['Biología','Química','Medio ambiente','Física']
sociales_6_7  = ['Historia', 'Geografía', 'Participación política','Filosofía']
lenguaje_6_7  = ['Comunicación y sistemas simbólicos','Producción e interpretación de textos']
matemati_6_7  = ['Aritmética','Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

######################################################################################################################################

asignaturas_8_9= ['Biología','Química','Medio ambiente','Física',
                  'Historia', 'Geografía', 'Participación política','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Álgebra', 'Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

ciencias_8_9  = ['Biología','Química','Medio ambiente','Física']
sociales_8_9  = ['Historia', 'Geografía', 'Participación política','Filosofía']
lenguaje_8_9  = ['Comunicación y sistemas simbólicos','Producción e interpretación de textos']
matemati_8_9  = ['Álgebra', 'Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

######################################################################################################################################


asignaturas_10=  ['Biología','Química','Medio ambiente','Física',
                  'Ciencias económicas', 'Ciencias políticas','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Trigonometría', 'Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']

ciencias_10  = ['Biología','Química','Medio ambiente','Física']
sociales_10  = ['Ciencias económicas', 'Ciencias políticas','Filosofía']
lenguaje_10  = ['Comunicación y sistemas simbólicos','Producción e interpretación de textos','Metodología']
matemati_10  = ['Trigonometría', 'Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']

######################################################################################################################################
asignaturas_11=  ['Química','Medio ambiente','Física',
                  'Ciencias económicas', 'Ciencias políticas','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Cálculo','Animaplanos', 'Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']

ciencias_11  = ['Química','Medio ambiente','Física']
sociales_11  = ['Ciencias económicas', 'Ciencias políticas','Filosofía']
lenguaje_11  = ['Comunicación y sistemas simbólicos','Producción e interpretación de textos','Metodología']
matemati_11  = ['Cálculo','Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']

vectores = {}
bloques = ['A','B','C','D']

for i in range(len(listado)):
    estudiante = listado.iloc[i, 0]
    grado = listado.iloc[i, 2]
    vectores[estudiante] = [''] * 10
    
    if grado in ['5','6','7']:
        for k in range(10):
            for bloque in bloques:
                notas_estudiante_actual = base[
                        (base.iloc[:, 2] == estudiante) &
                        (base.iloc[:, 3] == grado) &
                        (base.iloc[:, 6] == bloque) & 
                        (base.iloc[:, 5].isin(asignaturas_6_7))]
                    
                T = len(notas_estudiante_actual)
                if T == 75:
                    continue
                    
                if T < 75:
                    notas_c = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(ciencias_6_7)]
                    notas_s = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(sociales_6_7)]
                    notas_l = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(lenguaje_6_7)]
                    notas_m = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(matemati_6_7)]
                    
                    porcentaje_c = (len(notas_c) * 100) / 20
                    porcentaje_s = (len(notas_s) * 100) / 20
                    porcentaje_l = (len(notas_l) * 100) / 10
                    porcentaje_m = (len(notas_m) * 100) / 25
                        
                    menor_valor = min(porcentaje_c, porcentaje_s, porcentaje_l, porcentaje_m)
                        
                    if menor_valor == porcentaje_c:
                        vectores[estudiante][k] = 'C1'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Química', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_s:
                        vectores[estudiante][k] = 'S1'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Filosofía', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_l:
                        vectores[estudiante][k] = 'L'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Comunicación y sistemas simbólicos', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_m:
                        vectores[estudiante][k] = 'M1'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Estadística', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break

    if grado in ['8','9']:
        for k in range(10):
            for bloque in bloques:
                notas_estudiante_actual = base[
                        (base.iloc[:, 2] == estudiante) &
                        (base.iloc[:, 3] == grado) &
                        (base.iloc[:, 6] == bloque) & 
                        (base.iloc[:, 5].isin(asignaturas_8_9))]
                    
                T = len(notas_estudiante_actual)
                if T == 75:
                    continue
                    
                if T < 75:
                    notas_c = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(ciencias_8_9)]
                    notas_s = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(sociales_8_9)]
                    notas_l = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(lenguaje_8_9)]
                    notas_m = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(matemati_8_9)]
                    
                    porcentaje_c = (len(notas_c) * 100) / 20
                    porcentaje_s = (len(notas_s) * 100) / 20
                    porcentaje_l = (len(notas_l) * 100) / 10
                    porcentaje_m = (len(notas_m) * 100) / 25
                        
                    menor_valor = min(porcentaje_c, porcentaje_s, porcentaje_l, porcentaje_m)
              
                    if menor_valor == porcentaje_c:
                        if grado == '8':
                            vectores[estudiante][k] = 'C1'
                        if grado == '9':
                            vectores[estudiante][k] = 'C2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Química', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_s:
                        if grado == '8':
                            vectores[estudiante][k] = 'S1'
                        if grado == '9':
                            vectores[estudiante][k] = 'S2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Filosofía', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_l:
                        vectores[estudiante][k] = 'L'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Comunicación y sistemas simbólicos', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_m:
                        if grado == '8':
                            vectores[estudiante][k] = 'M1'
                        if grado == '9':
                            vectores[estudiante][k] = 'M2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Estadística', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break

    if grado in ['10']:
        for k in range(10):
            for bloque in bloques:
                notas_estudiante_actual = base[
                        (base.iloc[:, 2] == estudiante) &
                        (base.iloc[:, 3] == grado) &
                        (base.iloc[:, 6] == bloque) & 
                        (base.iloc[:, 5].isin(asignaturas_10))]
                    
                T = len(notas_estudiante_actual)
                if T == 75:
                    continue
                    
                if T < 75:
                    notas_c = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(ciencias_10)]
                    notas_s = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(sociales_10)]
                    notas_l = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(lenguaje_10)]
                    notas_m = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(matemati_10)]
                    
                    porcentaje_c = (len(notas_c) * 100) / 20
                    porcentaje_s = (len(notas_s) * 100) / 15
                    porcentaje_l = (len(notas_l) * 100) / 10
                    porcentaje_m = (len(notas_m) * 100) / 25
                        
                    menor_valor = min(porcentaje_c, porcentaje_s, porcentaje_l, porcentaje_m)
                        
                    if menor_valor == porcentaje_c:
                        vectores[estudiante][k] = 'C2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Química', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_s:
                        vectores[estudiante][k] = 'S2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Filosofía', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_l:
                        vectores[estudiante][k] = 'L'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Comunicación y sistemas simbólicos', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_m:
                        vectores[estudiante][k] = 'M2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Estadística', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break

    if grado in ['11']:
        for k in range(10):
            for bloque in bloques:
                notas_estudiante_actual = base[
                        (base.iloc[:, 2] == estudiante) &
                        (base.iloc[:, 3] == grado) &
                        (base.iloc[:, 6] == bloque) & 
                        (base.iloc[:, 5].isin(asignaturas_11))]
                    
                T = len(notas_estudiante_actual)
                if T == 70:
                    continue
                    
                if T < 70:
                    notas_c = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(ciencias_11)]
                    notas_s = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(sociales_11)]
                    notas_l = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(lenguaje_11)]
                    notas_m = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(matemati_11)]
                    
                    porcentaje_c = (len(notas_c) * 100) / 15
                    porcentaje_s = (len(notas_s) * 100) / 15
                    porcentaje_l = (len(notas_l) * 100) / 10
                    porcentaje_m = (len(notas_m) * 100) / 25
                        
                    menor_valor = min(porcentaje_c, porcentaje_s, porcentaje_l, porcentaje_m)
                        
                    if menor_valor == porcentaje_c:
                        vectores[estudiante][k] = 'C2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Química', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_s:
                        vectores[estudiante][k] = 'S2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Filosofía', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_l:
                        vectores[estudiante][k] = 'L'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Comunicación y sistemas simbólicos', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_m:
                        vectores[estudiante][k] = 'M2'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Estadística', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break

 

lista_de_vectores = [vectores[estudiante] for estudiante in vectores]
horario = pd.DataFrame(lista_de_vectores, index=vectores.keys())
#horario.to_excel('C:/Users/Admin/Desktop/GK/Horario/Areas bachillerato.xlsx')     

cupos1 = pd.read_excel('C:/Users/Admin/Desktop/GK/Horario/base planeacion.xlsx',sheet_name= 'cuposb', header= None)
cupos1 = cupos1.iloc[:7, :10]
cupos = pd.DataFrame(np.zeros((7, 10)))
cupos.iloc[:cupos1.shape[0], :cupos1.shape[1]] = cupos1.values
cupos = cupos.fillna(0)



# categorias_validas = set()

# for col in base_planeacion:
#    categorias_validas.update(base_planeacion[col].unique())

# categorias_validas = list(categorias_validas)


categorias_validas = ['C1','C2', 'S1','S2', 'L', 'M1','M2','E1','E2','V','LIC-C1','LIC-C2','LIC-S1','LIC-S2','LIC-L','LIC-M2','LIC-M1','PAZ','SIMULACRO-M1','SIMULACRO-C1','SIMULACRO-S1','SIMULACRO-L', 'SIMULACRO-E2','SIMULACRO-S2','EXTERNADO','CREA','KAIPOREMUN'                  
                      'VC','VS','VL','VM','VE','PG-C','PG-S1','PG-E1','PG-TS','PG-S2','PG-L','PG-M1','PG-M2','JP','PRUEBA','COMETAS','KPMUN','SIMONU','NO','FESTIVO','HALLOWEEN','UNIVERSIDADES','FES','TIVO','SUSTEN','TACIÓN','DESPEDIDA','VA','CA','CIO','NES',
                      'SABER-M2','SABER-C2','SABER-L','SABER-S2','PG','APOYOS','KAIPOMUN','MALOKA','FEST','IVO','PENSAR','FILBO','COUNTRY TOUR','MONSERRATE','OLIMPIADAS'
                      ]

for col in base_planeacion.columns[2:11]:  
    base_planeacion[col] = pd.Categorical(base_planeacion[col], categories=categorias_validas, ordered=False)

for i in range(len(base_planeacion)):
    estudiante = base_planeacion.iloc[i, 0]
    areas = vectores[estudiante]
    for k in range(10):
        area_actual = areas[k]
        
        if not(pd.isna(base_planeacion.iloc[i, k+2])):
            vectores[estudiante] = [vectores[estudiante][9]] + vectores[estudiante][:-1]
            continue

        for j in range(10):
            if area_actual == 'C1' and cupos.iloc[0, k] < 12:
                base_planeacion.iloc[i, k+2] = 'C1'
                cupos.iloc[0, k] += 1
                break
            elif area_actual == 'C2' and cupos.iloc[1, k] < 12:
                base_planeacion.iloc[i, k+2] = 'C2'
                cupos.iloc[1, k] += 1
                break
            elif area_actual == 'S1' and cupos.iloc[2, k] < 12:
                base_planeacion.iloc[i, k+2] = 'S1'
                cupos.iloc[2, k] += 1
                break
            elif area_actual == 'S2' and cupos.iloc[3, k] < 12:
                base_planeacion.iloc[i, k+2] = 'S2'
                cupos.iloc[3, k] += 1
                break
            elif area_actual == 'L' and cupos.iloc[4, k] < 12:
                base_planeacion.iloc[i, k+2] = 'L'
                cupos.iloc[4, k] += 1
                break
            elif area_actual == 'M1' and cupos.iloc[5, k] < 12:
                base_planeacion.iloc[i, k+2] = 'M1'
                cupos.iloc[5, k] += 1
                break
            elif area_actual == 'M2' and cupos.iloc[6, k] < 12:
                base_planeacion.iloc[i, k+2] = 'M2'
                cupos.iloc[6, k] += 1
                break
            else:
                area_actual = areas[(k + j + 1) % 10]

#base_planeacion.to_excel('C:/Users/Admin/Desktop/GK/Horario/Planeación bachillerato con huecos.xlsx', index = False)

for i in range(len(base_planeacion)):
    for k in range(10):
        if (pd.isna(base_planeacion.iloc[i, k+2])):
            columna = cupos.iloc[:,k]
            lista_cupos = columna.tolist()
            lista_cupos.sort()
            fila_estudiante_actual = listado[listado.iloc[:, 0] == base_planeacion.iloc[i, 0]]
            grado = fila_estudiante_actual.iloc[0,2]
            for h in range(7):
                if lista_cupos[h] == cupos.iloc[0,k] and cupos.iloc[0, k] < 12 and grado in ['6','7','8']:
                    base_planeacion.iloc[i,k+2] = 'C1'
                    cupos.iloc[0,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[1,k] and cupos.iloc[1, k] < 12 and grado in ['9','10','11']:
                    base_planeacion.iloc[i,k+2] = 'C2'
                    cupos.iloc[1,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[2,k] and cupos.iloc[2, k] < 12 and grado in ['6','7','8']: 
                    base_planeacion.iloc[i,k+2] = 'S1'
                    cupos.iloc[2,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[3,k] and cupos.iloc[3, k] < 12 and grado in ['9','10','11']:
                    base_planeacion.iloc[i,k+2] = 'S2'
                    cupos.iloc[3,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[4,k] and cupos.iloc[4, k] < 12:
                    base_planeacion.iloc[i,k+2] = 'L'
                    cupos.iloc[4,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[5,k] and cupos.iloc[5, k] < 12 and grado in ['6','7','8']:
                    base_planeacion.iloc[i,k+2] = 'M1'
                    cupos.iloc[5,k] += 1
                    break
                if lista_cupos[h] == cupos.iloc[6,k] and cupos.iloc[6, k] < 12 and grado in ['9','10','11']:
                    base_planeacion.iloc[i,k+2] = 'M2'
                    cupos.iloc[6,k] += 1
                    break

for i in range(len(base_planeacion)):
    for k in range(10):
        if (pd.isna(base_planeacion.iloc[i, k+2])):
            columna = cupos.iloc[:,k]
            lista_cupos = columna.tolist()
            lista_cupos.sort()
            fila_estudiante_actual = listado[listado.iloc[:, 0] == base_planeacion.iloc[i, 0]]
            if lista_cupos[0] == cupos.iloc[0,k]:
                base_planeacion.iloc[i,k+2] = 'C1'
                cupos.iloc[0,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[1,k]:
                base_planeacion.iloc[i,k+2] = 'C2'
                cupos.iloc[1,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[2,k]: 
                base_planeacion.iloc[i,k+2] = 'S1'
                cupos.iloc[2,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[3,k]:
                base_planeacion.iloc[i,k+2] = 'S2'
                cupos.iloc[3,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[4,k]:
                base_planeacion.iloc[i,k+2] = 'L'
                cupos.iloc[4,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[5,k]:
                base_planeacion.iloc[i,k+2] = 'M1'
                cupos.iloc[5,k] += 1
                break
            if lista_cupos[0] == cupos.iloc[6,k]:
                base_planeacion.iloc[i,k+2] = 'M2'
                cupos.iloc[6,k] += 1
                break

base_planeacion.to_excel('C:/Users/Admin/Desktop/GK/Horario/Planeación bachillerato.xlsx', index = False)