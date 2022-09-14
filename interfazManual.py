from csv import writer
from optparse import Values
from statistics import variance
from webbrowser import get
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import os
import sys
from subprocess import CREATE_NO_WINDOW
from Google import Create_Service

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def interf():
    global pantalla
    pantalla = tk.Tk()
    pantalla.geometry("1000x600")
    pantalla.title("Interfaz De Asignación de Licitaciones")
    path = resource_path("icono.ico")
    pantalla.iconbitmap(path)
    path = resource_path("logo.png")
    image = PhotoImage(file=path)
    image = image.subsample(2, 2)
    label = Label(image=image)
    label.pack()

#INSERTA URL
    Label(text="APLICACION -- ASIGNACION DE LICITACION", bg="#26CBA9",
          fg="black", width="36", height="1", font=("calibri", 15)).pack()
    Label(text="").pack()
    global urlgo
    urlgo = StringVar()
    global url

    Label(pantalla, text="INGRESE URL DE LICITACION",
          font=("calibri", 15)).pack()
    url = Entry(pantalla, width=60, textvariable=urlgo)
    url.pack()
    Label(pantalla).pack()

#INSERTA NOMBRE DEL RESPONSABLE
    Label(text="SELECCIONE RESPONSABLE LICITACION", bg="#26CBA9",
          fg="black", width="36", height="1", font=("calibri", 15)).pack()
    Label(text="").pack()
    global respons
    respons = StringVar()

# CREO LA LISTA Y EL MENU DESPLEGABLE
    OptionList = [" ", "Martin Chacon", "Gerko Marfull"]
    OptionList.sort()
    global variable
    variable = tk.StringVar(pantalla)
    variable.set(OptionList[0])
    opt = tk.OptionMenu(pantalla, variable, *OptionList)
    opt.config(width=33, font=('calibri', 15), background='White')
    opt.pack()
    opt = tk.OptionMenu(pantalla, variable, *OptionList)
    Label(text="").pack

    #BOTON COPIAR AL PORTAPELES// NUEVAS FUNCIONES AGREGARLAS ARRIBA DE ESTE MODULO
    Button(pantalla, text="Enviar", command=copy, padx=155, pady=10,
           background='red', font=("calibri", 15)).pack(side=tk.BOTTOM)
    pantalla.mainloop()
    global nombre_asignado
    nombre_asignado = variable.get()


def validacion(idLicitacion):
    CLIENT_SECRET_FILE = os.path.join('client_secret.json')
    API_SERVICE_NAME = 'sheets'  # NOMBRE DE LA API
    API_VERSION = 'v4'  # VERSIÓN DE LA API
    # PERMISOS DE MODIFICACION DE GOOGLE SHEETS
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    gsheetId1 = '1fletI3lMA7R92AIjE4nYzcwcXsLOz2Ry1OhhiODS8eg'  # GERENCIA
    gsheetId2 = '1umHVdhFQ3Br06yj5JzWPTothXf9C1PcrVTPGrdeKcwk'  # PERSONAL
    global s
    s = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME,
                       API_VERSION, SCOPES)
    gs = s.spreadsheets()
    col1 = gs.values().get(spreadsheetId=gsheetId1, range='LICITACIONES!A2:A').execute()
    col2 = gs.values().get(spreadsheetId=gsheetId2, range='LICITACIONES!F8:F').execute()
    dataX1 = col1.get('values')
    dataX2 = col2.get('values')
    df1 = pd.DataFrame(dataX1)
    df2 = pd.DataFrame(dataX2)
    dfF = df2.append(df1, ignore_index=True)
    
    for idLc in dfF.index:
        if idLicitacion == str(dfF[0][idLc]):
            return False

def copy():
    options = webdriver.ChromeOptions()
    website = (urlgo.get())
    options.add_argument('headless')
    chrome_service = ChromeService(resource_path('chromedriver.exe'))
    chrome_service.creationflags = CREATE_NO_WINDOW
    driver = webdriver.Chrome(service=chrome_service, options=options)
    driver.get(website)
    fechaCierre = driver.find_element(By.ID, 'lblFicha3Cierre').text
    organismo = driver.find_element(By.ID, 'lblResponsable').text
    idLicitacion = driver.find_element(By.ID, 'lblNumLicitacion').text
    nombreLicitacion = driver.find_element(By.ID, 'lblNombreLicitacion').text
    tipos = driver.find_element(By.ID, 'lblFicha1Tipo').text
    tipo = tipos[-4:].replace("(", "").replace(")", "")
    validateToF = validacion(idLicitacion)

#MOSTRAMOS EL DATAFRAME DEL RESPONSABLE ASIGNADO
    lici = {
        "RESPONSABLE": [""], "NOMBRE LICITACIÓN": [nombreLicitacion], "FECHA DE CIERRE": [fechaCierre], "DIAS RESTANTES": ["=SI(B8='';'x';SI.ERROR(SIFECHA($J$3;C8;'D'); 'VENCIDA'))"], "ORGANISMO": [organismo], "ID LICITACION": [idLicitacion], "LINK LICITACION": [website], "LINK HOJA DE CALCULO": [""], "ESTADO ENVIO OFERTA": ["=SI(F8=\"\";\"\";IMPORTRANGE(F8;\"Plantilla!C10\"))"], "ESTADO ADJUDICACION OFERTA": ["=SI(F8=\"\";\"\";IMPORTRANGE(F8;\"Plantilla!E10\"))"], "EMPRESA LICITADORA": ["=SI(F8=\"\";\"\";IMPORTRANGE(F8;\"Plantilla!E11\"))"], "TIPOS DE LICITACION": [tipo], "OBSERVACION": ["MANUAL"]
    }
    df = pd.DataFrame(lici)

    #CODIGO EN DONDE SE LE DAN PARAMETROS A LA API: LAS KEY, NOMBRE DE API, ETC, PARA PODER OPERAR CON SUS COMANDOS.

    # DENTRO DE ESTE ARCHIVO JSON, SE ENCUENTRAN LAS KEYS.(YA ESTÁN CONFIGURADAS)
    CLIENT_SECRET_FILE = os.path.join('client_secret.json')
    API_SERVICE_NAME = 'sheets'  # NOMBRE DE LA API
    API_VERSION = 'v4'  # VERSIÓN DE LA API
    # PERMISOS DE MODIFICACION DE GOOGLE SHEETS
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    #CONEXIÓN CON LA API, CON "service", SE PUEDE UTILIZAR LOS ENDPOINT DE LA API, REALIZA LA CONEXION IMPORTANDO LA FUNCIÓN CREATE_SERVICE DE EL SCRIPT Google.py
    service = Create_Service(
        CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

    #EL ID DE LA HOJA DE CALCULO, ESTE SE PUEDE ENCONTRAR EN EL LINK DE LA HOJA DE CALCULO, entre "d/" y "/edit"
    #EJEMPLO: https://docs.google.com/spreadsheets/d/  --->1DatQ3q5h6xrtRDnOM3MhRWCDgbtVACy2FeZAxIqNbzY<---   /edit#gid=147158629
    spreadsheet_id = '10zw5w9fNuFHVJCZJFmORLsbFSxY2eqz2WIY9ajE3Ffw'
    sheet_id = '147158629'

    def ejecutar():
        #EJEMPLO DE UTILIZACION DE LA API, MAYOR DETALLE EN LISTA DE YOUTUBE
        request_body = {
            'requests': [
                {
                    'insertDimension': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'ROWS',
                            'startIndex': 7,
                            'endIndex': 8,
                        }
                    }
                }
            ]
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=request_body
        ).execute()

#CODIGO EN DONDE SE LE DAN PARAMETROS A LA API: LAS KEY, NOMBRE DE API, ETC, PARA PODER OPERAR CON SUS COMANDOS.
        mySpreadsheets = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id).execute()

        worksheet_name = 'LICITACIONES!'
        cell_range_insert = 'A7'
        values = (
            ('RESPONSABLE', 'NOMBRE LICITACIÓN', 'FECHA DE CIERRE', 'DIAS REST.', 'ORGANISMO', 'ID LICITACION', 'LINK LICITACIÓN', 'LINK HOJA DE CALCULO',
             'ESTADO ENVIO OFERTA', 'ESTADO ADJUDICACION OFERTA', 'EMPRESA LICITADORA', 'TIPO DE LICITACION', 'OBSERVACIÓN', 'TOMAR', 'DESCARTAR', 'HACER SÍ O SÍ'),
            (str(df.iat[0, 0]), str(df.iat[0, 1]), str(df.iat[0, 2]), str(df.iat[0, 3]), str(df.iat[0, 4]), str(df.iat[0, 5]),
             str(df.iat[0, 6]), str(df.iat[0, 7]), str(df.iat[0, 8]), str(
                 df.iat[0, 9]), str(df.iat[0, 10]), str(df.iat[0, 11]),
             str(df.iat[0, 12]))
        )
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': values
        }

        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            valueInputOption='USER_ENTERED',
            range=worksheet_name + cell_range_insert,
            body=value_range_body
        ).execute()

        messagebox.showinfo(title="Datos enviados correctamente",
                            message="Licitación enviada a la base de datos.")

    if validateToF == False:
        messagebox.showinfo(title="Datos enviados anteriormente",
                            message="La licitación NO se ha enviado a la base de datos porque ya existe.")
    else:
        ejecutar()

interf()
