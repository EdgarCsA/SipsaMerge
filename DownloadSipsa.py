import unicodecsv
import datetime
import xlrd
import csv
import xlwt
import openpyxl
import win32com.client
import os
from os import sys
from win32com.client import constants as c
import pandas as pd
import numpy as np

#%cd C:\Users\Edgar Céspedes\Documents\DescargarArchivos

def mesANumero(mes):
    m = {
        'enero': "01",
        'febrero': "02",
        'marzo': "03",
        'abril': "04",
        'mayo': "05",
        'junio': "06",
        'julio': "07",
        'agosto': "08",
        'septiembre': "09",
        'octubre': "10",
        'noviembre': "11",
        'diciembre': "12"
        }
    out= str(m[mes.lower()])
    return (out)

def AgregarProdu(general, nombres, p):
    ListaMeses=["enero","febrero","marzo","abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    jk=[]
    ji=[]
    jo=[]
    for mes2 in ListaMeses:
        for dia3 in range(1,31):
            dia=str(dia3)
            fileDir, fileName = os.path.split(mes2+"_"+dia+".xls")
           

            if (os.path.isfile(fileName)):
                wb = xlrd.open_workbook(mes2+"_"+dia+".xls")
                sh = wb.sheet_by_index(0)
                jo=sh.row_values(p)
                jk=jo[1:len(jo)]
                ji.append(jk[0])
             
                for i in range(1,int(len(jk)/2)):
                    #print(i)
                    ji.append(jk[i*2])  
    general.append(ji)
    nombres.append(jo[0])
    return general, nombres

ListaMeses=["enero","febrero","marzo","abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


AnodelDescarga=input("Ingrese el año del conjunto de datos que quiere descargar:")
for mes in ListaMeses:
    for dia2 in range(1,31):
        dia=str(dia2)
        dls = "http://www.dane.gov.co/files/investigaciones/agropecuario/sipsa/mayoristas_"+mes+"_"+dia+"_"+AnodelDescarga+".xls"
        print("Descargado->mayoristas_"+mes+"_"+dia+"_"+AnodelDescarga+".xls")
        import requests
        r=response = requests.get(dls)
        response.status_code
        if response.status_code==200:
            Tidoc=mes+"_"+dia+".xls"
            with open(Tidoc, "wb") as code:
                code.write(r.content)




general=[]
jk=[]
date=[]
for mes2 in ListaMeses:
    for dia3 in range(1,31):
        dia=str(dia3)
        fileDir, fileName = os.path.split(mes2+"_"+dia+".xls")
        
        if (os.path.isfile(fileName)):
            wb = xlrd.open_workbook(mes2+"_"+dia+".xls")
            sh = wb.sheet_by_index(0)
            mes=mesANumero(mes2)
            Fecha=dia+"/"+mes+"/"+AnodelDescarga     
            #for N in range(5,18):
            jo=sh.row_values(2)
            while '' in jo:
                jo.remove('')
            jk=jk+jo[1:len(jo)]
            #general.append(jk)
            for k in range(len(jo[1:len(jo)])):
                date.append(Fecha)
            #general.insert(0,date)
            jo.insert(0, Fecha)
            

general.append(jk)
general.insert(0,date)
nombres=["Fecha", "Mercado"]

import itertools as it
for i in it.chain(range(5, 18), range(20,36), range(38,43)):
    general2, nombres2= AgregarProdu(general, nombres, i)


dfBD1 = pd.DataFrame(general)
dfBD1= pd.DataFrame.transpose(dfBD1)


dfBD1.columns=nombres

dfBD1.to_csv('Final.csv')
print("El archivo de datos agregado es Final.csv")
