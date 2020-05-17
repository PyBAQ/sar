import json
import pandas as pd
from xlrd.biffh import XLRDError
from datetime import datetime, timedelta


# Se carga archivo de configuración
with open('config.json', 'r') as file:
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

# Renombrado de columnas para estandarizar

meetupList.columns = ['Nombre Completo', 'Direccion de correo electronico']

# Lectura de archivo de asistentes registrados en Google form
try:
    googleList = pd.read_excel(
        config['GOOGLE']['FILE_ROOT'],
        config['GOOGLE']['FILE_SHEET'],
        usecols=config['GOOGLE']['FILE_COLS'],
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
    print(
        "Error leyendo archivo",
        config['GOOGLE']['FILE_ROOT']
    )
else:
    print("Lectura de archivo de Google exitosa")


# Eliminando filas de división en dataframe de archivo de Google
googleList.dropna(inplace=True)

# Filtrando archivo de Google con inscripciones dentro de las fechas configuradas

publication_date = pd.to_datetime(
    config['EVENT']['PUBLICATION_DATE'])

closing_date = pd.to_datetime(
    config['EVENT']['CLOSING_DATE'])

mask = (googleList['Marca temporal'] >= publication_date) & (
    googleList['Marca temporal'] <= closing_date)

filteredGoogleList = googleList.loc[mask]

# Comparación de dataframes - recorrer listado de Meetup para buscar coincidencias por correo o nombre con listado fitrado de Google


def findInFilteredList(participanteMeetup):
    for indexG, participanteGoogle in filteredGoogleList.iterrows():
        if participanteMeetup['Direccion de correo electronico'] == participanteGoogle['Direccion de correo electronico']:
            print('Encontrado por correo: '+participanteMeetup['Nombre Completo'] +
                  ' Registro en Google Forms Ok')
            return True

        elif participanteMeetup['Nombre Completo'].lower() == participanteGoogle['Nombre Completo'].lower():
            print('Encontrado por nombre: ' + participanteMeetup['Nombre Completo'] +
                  ' Registro en Google Forms Ok')
            return True
    else:
        return False


# Comparación de dataframes - buscar datos de participantes de Meetup en archivo histórico de Google
participantsInHistorical = []


def findInHistoricalList(participanteMeetup):
    for indexH, participanteHistorico in googleList.iterrows():
        if participanteMeetup['Direccion de correo electronico'] == participanteHistorico['Direccion de correo electronico']:
            print('Encontrado por correo: ' +
                  participanteMeetup['Nombre Completo'] + ' Registro en Google Historico Ok')
            participantsInHistorical.append(dict(participanteHistorico))
            return True

        elif participanteMeetup['Nombre Completo'].lower() == participanteHistorico['Nombre Completo'].lower():
            print('Encontrado por nombre: ' + participanteMeetup['Nombre Completo'] +
                  ' Registro en Google Historico Ok')
            participantsInHistorical.append(dict(participanteHistorico))
            return True
    else:
        return False


for indexM, participanteMeetup in meetupList.iterrows():
    if not findInFilteredList(participanteMeetup):
        findInHistoricalList(participanteMeetup)

# Comparación de dataframes - generar dataframe resultante
participantsInHistoricalDF = pd.DataFrame(participantsInHistorical)
participantsList = pd.concat(
    [filteredGoogleList, participantsInHistoricalDF], sort=False)

# Organizando alfabeticamente por nombre para facilitar busqueda
participantsList.sort_values('Nombre Completo', inplace=True)

# Eliminando registros con correos duplicados
participantsList.drop_duplicates(
    subset='Direccion de correo electronico', keep='first', inplace=True)

# Ordenando columnas para archivo
participantsList = participantsList
[[
    'Nombre completo', 'Tipo de identificacion', 'Direccion de correo electronico', 'Telefono de contacto',
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
