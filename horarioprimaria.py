import pandas as pd # type: ignore
from corregir_nombres import corregir_nombre
import numpy as np # type: ignore

#leer base de datos
base = pd.read_excel('C:/Users/Admin/Desktop/GK/BASE DE DATOS 2025.xlsx', sheet_name = 'GK2025')
base['GRADO'] = base['GRADO'].astype(str)
base['ESTUDIANTE'] = base['ESTUDIANTE'].apply(corregir_nombre)

#leer listado general
prim = ['1','2','3','4','5'] #listado de numeros del 1 al 5
listado = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name = 'g') # lee el listado general
listado['GRADO'] = listado['GRADO'].astype(str) #convierte los datos de la columna de grado a cadenas de texto (ya no son numeros)
listado = listado[listado.iloc[:,2].isin(prim)] # filtra del listado general solo los estudiantes de grados 1 a 5
listado['ESTUDIANTE'] = listado['ESTUDIANTE'].apply(corregir_nombre)
base_planeacion = pd.read_excel('C:/Users/Admin/Desktop/GK/Horario/base planeacion.xlsx', sheet_name = 'primaria') #Leer base planeacion
base_planeacion['Estudiante'] = base_planeacion['Estudiante'].apply(corregir_nombre)


#En este bloque se organizan los ordenes del grupo de ingles segun la letra que se ingrese para luego ordenar la 

a = input('INGRESE EL GRUPO DE INGLES :')

if a == 'A':
    orden = ['A','B','C','D','E']
if a == 'B':
    orden = ['B','C','D','E','A']
if a == 'C':
    orden = ['C','D','E','A','B']
if a == 'D':
    orden = ['D','E','A','B','C']
if a == 'E':
    orden = ['E','A','B','C','D']


base_planeacion['INGLÉS'] = pd.Categorical(base_planeacion['INGLÉS'], categories=orden, ordered=True)
base_planeacion = base_planeacion.sort_values('INGLÉS')
base_planeacion = base_planeacion.reset_index(drop=True)

asignaturas_primaria= ['Biología','Química','Medio ambiente','Física',
                       'Historia', 'Geografía', 'Participación política',
                       'Comunicación y sistemas simbólicos','Producción e interpretación de textos','Pensamiento religioso',
                       'Aritmética','Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

ciencias_primaria  = ['Biología','Química','Física','Medio ambiente']
sociales_primaria  = ['Historia', 'Geografía', 'Participación política']
lenguaje_primaria  = ['Comunicación y sistemas simbólicos','Producción e interpretación de textos','Pensamiento religioso']
matemati_primaria  = ['Aritmética','Estadística', 'Geometría','Dibujo técnico','Sistemas']


vectores = {}
bloques = ['A','B','C','D']

for i in range(len(listado)):
    estudiante = listado.iloc[i, 0]
    grado = listado.iloc[i, 2]
    vectores[estudiante] = [''] * 10
    
    if grado in ['1','2','3','4','5']:
        for k in range(10):
            for bloque in bloques:
                notas_estudiante_actual = base[
                        (base.iloc[:, 2] == estudiante) &
                        (base.iloc[:, 3] == grado) &
                        (base.iloc[:, 6] == bloque) & 
                        (base.iloc[:, 5].isin(asignaturas_primaria))]
                    
                T = len(notas_estudiante_actual)
                if T == 75:
                    continue
                    
                if T < 75:
                    notas_c = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(ciencias_primaria)]
                    notas_s = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(sociales_primaria)]
                    notas_l = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(lenguaje_primaria)]
                    notas_m = notas_estudiante_actual[notas_estudiante_actual['ASIGNATURA'].isin(matemati_primaria)]
                    
                    porcentaje_c = (len(notas_c) * 100) / 20
                    porcentaje_s = (len(notas_s) * 100) / 15
                    porcentaje_l = (len(notas_l) * 100) / 15
                    porcentaje_m = (len(notas_m) * 100) / 25
                        
                    menor_valor = min(porcentaje_c, porcentaje_s, porcentaje_l, porcentaje_m)
                        
                    if menor_valor == porcentaje_c:
                        vectores[estudiante][k] = 'C'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Química', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break  
                    elif menor_valor == porcentaje_s:
                        vectores[estudiante][k] = 'S'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Historia', f'{bloque}', 'X', 'X', 'X', 'X']
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
                        vectores[estudiante][k] = 'M'
                        lista_fila = ['X', 'X', f'{estudiante}', f'{grado}', 'X', 'Estadística', f'{bloque}', 'X', 'X', 'X', 'X']
                        fila_para_agregar = pd.DataFrame([lista_fila], columns=base.columns)
                        base = pd.concat([base, fila_para_agregar], ignore_index=True)
                        break
                  
lista_de_vectores = [vectores[estudiante] for estudiante in vectores]
horario = pd.DataFrame(lista_de_vectores, index=vectores.keys())
#horario.to_excel('C:/Users/Admin/Desktop/GK/Horario/Areas primaria.xlsx')     


cupos1 = pd.read_excel('C:/Users/Admin/Desktop/GK/Horario/base planeacion.xlsx',sheet_name= 'cuposp', header= None)
cupos1 = cupos1.iloc[:4, :10]
cupos = pd.DataFrame(np.zeros((4, 10)))
cupos.iloc[:cupos1.shape[0], :cupos1.shape[1]] = cupos1.values
cupos = cupos.fillna(0)

categorias_validas = ['C', 'S', 'L', 'M','E','LIC-C','LIC-S','LIC-L','LIC-M','E-STEAM','FESTIVO','JP','CLUB','STEM','PRUEBA','ANIMAPLANOS','PENSAR','FILBO','SIMONU','COUNTRY TOUR',
                      'COMETAS','PAZ','KPMUN','CREA','OLIMPIADAS','HALLOWEEN','FES','TIVO','DESPEDIDA','VA','CA','CIO','NES','KAIPOMUN','MALOKA','FEST','IVO','MONSERRATE']

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
            if area_actual == 'C' and cupos.iloc[0, k] < 11:
                base_planeacion.iloc[i, k+2] = 'C'
                cupos.iloc[0, k] += 1
                break
            elif area_actual == 'S' and cupos.iloc[1, k] < 11:
                base_planeacion.iloc[i, k+2] = 'S'
                cupos.iloc[1, k] += 1
                break
            elif area_actual == 'L' and cupos.iloc[2, k] < 11:
                base_planeacion.iloc[i, k+2] = 'L'
                cupos.iloc[2, k] += 1
                break
            elif area_actual == 'M' and cupos.iloc[3, k] < 11:
                base_planeacion.iloc[i, k+2] = 'M'
                cupos.iloc[3, k] += 1
                break
            else:
                area_actual = areas[(k + j + 1) % 10]

#base_planeacion.to_excel('C:/Users/Admin/Desktop/GK/Horario/Planeación primaria con huecos.xlsx', index = False)

for i in range(len(base_planeacion)):
    for k in range(10):
        if (pd.isna(base_planeacion.iloc[i, k+2])):
            columna = cupos.iloc[:,k]
            lista_cupos = columna.tolist()
            mayor_cupo = min(lista_cupos)
            if mayor_cupo == cupos.iloc[0,k]:
                base_planeacion.iloc[i,k+2] = 'C'
                cupos.iloc[0,k] += 1
                continue
            if mayor_cupo == cupos.iloc[1,k]:
                base_planeacion.iloc[i,k+2] = 'S'
                cupos.iloc[1,k] += 1
                continue
            if mayor_cupo == cupos.iloc[2,k]:
                base_planeacion.iloc[i,k+2] = 'L'
                cupos.iloc[2,k] += 1
                continue
            if mayor_cupo == cupos.iloc[3,k]:
                base_planeacion.iloc[i,k+2] = 'M'
                cupos.iloc[3,k] += 1
                continue

base_planeacion.to_excel('C:/Users/Admin/Desktop/GK/Horario/Planeación primaria.xlsx', index = False)