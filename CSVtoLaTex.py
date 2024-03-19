import pandas as pd
from tabulate import tabulate
import subprocess

# Supongamos que tienes el siguiente DataFrame
'''
df = pd.DataFrame({
    'Nombre': ['Ana', 'Juan', 'Carlos'],
    'Edad': [25, 32, 18],
    'Ciudad': ['Sevilla', 'Madrid', 'Barcelona']
})
'''



def creaInforme(nombre_hoja):

    df = pd.read_excel('Resultados_CORTO.xlsx', sheet_name=nombre_hoja)


    # Imprime los datos de la columna
    print(df)

    numNodo = df.iat[1,1]
    numCortos = df.iat[3,1]
    tiempoCorto = df.iat[5,1]
    tiempoInicial = df.iat[7,1]
    tiempoFinal = df.iat[9,1]

    '''
        Obtención del incio y del final de los cortos 
    '''


    iniciosCorto = []
    finalesCorto = []
    inicioSTR =[]
    finSTR =[]
    Duracion = []
    longitudInicios = len(df['Unnamed: 2'])
    longitudFinales = len(df['Unnamed: 3'])

    for i in range(longitudInicios) :
        if(not(pd.isnull(df.iat[i,2]) ) and df.iat[i,2] != "Tiempo Inicio Corto"):
            iniciosCorto.append(df.iat[i,2]) #.strftime("%Y-%m-%d %H:%M:%S")


    for i in range(longitudFinales) :
        if(not(pd.isnull(df.iat[i,3]) ) and df.iat[i,3] != "Tiempo Fin Corto"):
            finalesCorto.append(df.iat[i,3])


    for i in range(len(finalesCorto)): 
        Duracion.append(str(finalesCorto[i]-iniciosCorto[i])) 

    print(Duracion)

    for tiempos in iniciosCorto:
        inicioSTR.append(tiempos.strftime("%Y-%m-%d %H:%M:%S"))

    for tiempos in finalesCorto:
        finSTR.append(tiempos.strftime("%Y-%m-%d %H:%M:%S"))

    # Crea una lista de listas (cada lista interna es una columna de la tabla)
    tabla = list(zip(inicioSTR, finSTR, Duracion))

    # Nombres de las columnas
    nombres_columnas = ['Número de Cortocircuito', 'Inicio', 'Fin', 'Tiempo']

    # Crea la tabla de LaTeX
    tabla_latex = tabulate(tabla, headers=nombres_columnas, tablefmt="latex", showindex='always')
    # Crear el documento de LaTeX
    documento_latex = f"""
    \\documentclass{{article}}
    \\usepackage[utf8]{{inputenc}}
    \\usepackage{{booktabs}}


    \\begin{{document}}

    \\section {{{numNodo}}}
    \\subsection {{Resumen}} 

    El nodo tuvo un total de {numCortos} cortocircuitos comprendidos desde {tiempoInicial} hasta {tiempoFinal}, haciendo un total de {tiempoCorto} horas en cortocircuitos.


    {tabla_latex}

    \\end{{document}}
    """
    doc = nombre_hoja + '.tex'
    # Guardar el documento de LaTeX en un archivo .tex
    with open(doc, 'w') as f:
        f.write(documento_latex)

    subprocess.call(["pdflatex", doc])



# Cargar el archivo Excel
excel_file = pd.ExcelFile('Resultados_CORTO.xlsx')

# Obtener los nombres de las hojas
sheet_names = excel_file.sheet_names
sheet_names.pop(0)
print("Nombres de las hojas disponibles:")
for hoja in sheet_names:
    creaInforme(hoja)
    

