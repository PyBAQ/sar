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
    print("Lectura exitosa")


# Lectura de archivo de asistentes registrados en Google form


# Comparación de dataframes
