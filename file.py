import json
import pandas as pd
from xlrd.biffh import XLRDError


# Se carga archivo de configuración
with open('config.json','r') as file:
    config = json.load(file)


# Lectura de archivo de asistentes registrados en Meetup
try:
    meetupList = pd.read_excel(
        config['MEETUP']['FILE_ROOT'],
        config['MEETUP']['FILE_SHEET'],
        usecols=config['MEETUP']['FILE_COLS']
    )
    if meetupList.empty:
        print("El archivo de participantes de Meetup está vacío")
        exit
except PermissionError:
    print("No se tienen permisos para acceder en la ruta especificada")
except FileNotFoundError:
    print(
        "No se encontró el archivo en la ruta:",
        config['MEETUP']['FILE_ROOT']
    )
except XLRDError as xlrd:
    print("Error leyendo archivo", config['MEETUP']['FILE_ROOT'])
else:
    print("Lectura de archivo de Meetup exitosa")


# Lectura de archivo de asistentes registrados en Google form
try:
    googleList = pd.read_excel(
        config['GOOGLE']['FILE_ROOT'],
        config['GOOGLE']['FILE_SHEET'],
        usecols = config['GOOGLE']['FILE_COLS'],
        dtype=config['GOOGLE']['FILE_COLS_FORMAT']
    )
    if googleList.empty:
        print("El archivo de participantes de Google Form está vacío")
        exit
except PermissionError:
    print("No se tienen permisos para acceder en la ruta especificada")
except FileNotFoundError:
    print(
        "No se encontró el archivo en la ruta:",
        config['GOOGLE']['FILE_ROOT']
    )
except XLRDError as xlrd:
    print (
        "Error leyendo archivo",
        config['GOOGLE']['FILE_ROOT']
    )
else:
    print("Lectura de archivo de Google exitosa")


# Eliminando filas de división en dataframe de archivo de Google
googleList.dropna(inplace = True)


# Comparación de dataframes - reemplazar participantsList por el dataframe resultante
participantsList=googleList

# Ordenando columnas para archivo
participantsList = participantsList
[[
    'Nombre completo', 'Tipo de identificacion',
    'Direccion de correo electronico', 'Telefono de contacto',
    'Empresa y/o profesión y/o actividad'
]]

# Escribiendo resultados en nuevo archivo
try:
    if participantsList.empty:
        print("El archivo de participantes de Meetup está vacío")
    else:
        with pd.ExcelWriter(config['PARTICIPANTS']['FILE_ROOT']) as writer:
            participantsList.to_excel(writer, index=False)
            writer.save()
except PermissionError:
    print("No se tienen permisos para acceder en la ruta especificada")
else:
    print("Creación de archivo de asistentes exitosa")
    