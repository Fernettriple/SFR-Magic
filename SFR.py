#!/usr/bin/env python3

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
import pandas as pd
import win32com.client as win32
import os
import numpy as np

###
#TODO VEMOS QUE ONDA ESTOoo
#MENSAJE DE BIENVENIDA, EXPLICAR QUE LOS NOMBRES DE LOS REPORTES TIENEN QUE ESTAR EN MAYUSCULAS Y ESO

#Esto me permite convertir la wea a .xlsx. lo saque de StackOverflow (NO CONSERVA H-LINKS)
#Hago lo mismo para los dos reportes de IP
# if 'IP RETURN.xlsx' not in os.listdir('.'):
#     if 'IP RETURN.xls' in os.listdir('.'):
#         fname = os.getcwd()+"\\IP RETURN.xls"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#     wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
#     wb.Close()                               #FileFormat = 56 is for .xls extension
#     excel.Application.Quit()

# if 'IP SHIPMENT.xlsx' not in os.listdir('.'):
#     if 'IP SHIPMENT.xls' in os.listdir('.'):
#         fname = os.getcwd()+"\\IP SHIPMENT.xls"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#     wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
#     wb.Close()                               #FileFormat = 56 is for .xls extension
#     excel.Application.Quit()


#Abro el archivo REPORT.xlsx para sacer la info necesaria
wb = openpyxl.load_workbook('REPORT.xlsx')    
ws=wb['Report']


#Extraigo el numero del sitio asi despues hago magia
Numero_de_sitio=ws["M2"].value
Nombre_de_archivo=str(Numero_de_sitio)+' COV Site File Review.xlsx'

# Voy copiando las columnas que necesito
Excel={}
Encabezados=['Level','Category','Document Type','Ref Model Subtype','Study Item Name','Ref Model ID','Site Personnel Name','Document Date','Expiration Date']
Max_Column=ws.max_column

# Busco usando la lista de encabezados, las coordenadas
for Encabezados_de_colunmnas in Encabezados:
    Excel[Encabezados_de_colunmnas]=[]
    for col_num in range (1,Max_Column):
        if ws.cell(row=1,column=col_num).value==Encabezados_de_colunmnas:            
            Letra_de_columna=get_column_letter(col_num)
            if ws.cell(row=1,column=col_num).value=='Document Date' or ws.cell(row=1,column=col_num).value=='Expiration Date':
                for cell in ws[Letra_de_columna]:
                        try: #los que son "None"
                            cell.value=cell.value.strftime('%d-%b-%Y')
                        except:
                            continue
            for cell in ws[Letra_de_columna]:
                Excel[Encabezados_de_colunmnas].append(cell.value)
# Con esto ya saque toda la info del REPORT.xlsx y lo tengo en el diccionario "Excel"

#Abro el archivo TEMPLATE.xlsx para pegar la info sacada de REPORT.xlsx.

wb = openpyxl.load_workbook('TEMPLATE.xlsx')    
ws=wb['Site']
Max_Row=ws.max_row
Max_Column=ws.max_column

# Antes de cualquier cosa, los encabezados del template y del report no son iguales, asi que hago la conversion asi no hay problema

Excel['Class']=Excel['Level']
Excel['Document Category']=Excel['Category']
Excel['Document Name']=Excel['Study Item Name']
Excel['Site personnel name']=Excel['Site Personnel Name']
Excel['Document date']=Excel['Document Date']
Excel['Expiration date']=Excel['Expiration Date']
# Una vez esto hecho, a agregar el contenido de esas listas al fondo de cada columna
# Parseo el diccionario vs los encabezados

for key in Excel.keys():
    for column_num in range (1,Max_Column): #Reviso todos los encabezados
        if key==ws.cell(1,column_num).value: #Si el encabezado coincide con la key que tengo
            for row_num in range (1,len(Excel[key])): 
                ws.cell(column=column_num,row=Max_Row+row_num-2).value=Excel[key][row_num] #Agrego todos los valores de la lista al final del ws

#Guardo todo antes del procesador

wb.save(Nombre_de_archivo)

#TODO
#Procesado
#Abro el documento con Pandas

filename=os.getcwd()+Nombre_de_archivo #consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa

SFR= pd.read_excel(filename, sheet_name='Site',header=0)
#Escanear nombres en CV y SubI de los CVs para predecir el study team. Luego, pedir input del DOA y predecir q se tiene q pedir en base a lo que hay. se podra usar un CTMS report?
#Buscar ultima FDA1572 y si es de hace dos años, preguntar si es la ultima.
#usando el reporte de visitas, checkear que estan todas las cfm, fup, svr
#los temp/CALIBRATION logs lo puedo checkear desde los reporte de IP vs IP RETURNED
#si es local o central tmb lo puedo sacar del log 





##IP SHIPMENTS/PL AND RETURNS


#IP Shipment confirmation
filename=os.getcwd()+"\\IP SHIPMENT.xlsx" #consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa
df= pd.read_excel(filename, sheet_name='Sheet',header=2) #Importo el excel en pandas. Header=2 porq ahi estan los encabezados
for index,row in df['Ship to Site Number'].iteritems(): #checkeo cada row de la columna "ship to site numbers"
    if row!=Numero_de_sitio:
        df.loc[index,'Ship to Site Number']=np.nan #Si la row no es del sitio q me interesa, le mando np.nan
    if df.loc[index,'Shipment Status']!='Received': 
        df.loc[index,'Shipment Status']=np.nan #Si el shipment no fue recibido, le mando np.nan
       
df.dropna(subset=['Ship to Site Number','Shipment Status'],inplace=True) #ahora, dropeo todos los np.nan, asi me quedo solo con los shipments recibidos y de mi sitio

#Esto... no me acuerdo porq pero messirve. Creo una DF solo de la info que me importa con las columnas q me importan
cols=list(df.columns) 
IP_SHIPMENT=df[[cols[0]]+[cols[8]]+[cols[9]]] #0= Shipment Number, 8 = Shipped Date, 9= Received Date

#IP Return
filename=os.getcwd()+"\\IP RETURN.xlsx"#consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa
df= pd.read_excel(filename, sheet_name='Sheet',header=2)#Importo el excel en pandas. Header=2 porq ahi estan los encabezados
for index,row in df['Ship from Site Number'].iteritems(): #checkeo cada row de la columna "ship to site numbers"
    if row!=Numero_de_sitio:
        df.loc[index,'Ship from Site Number']=np.nan #Si la row no es del sitio q me interesa, le mando np.nan
    if df.loc[index,'Return Shipment Status']!='Received':
        df.loc[index,'Return Shipment Status']=np.nan #Si el shipment no fue recibido, le mando np.nan
       
df.dropna(subset=['Ship from Site Number','Return Shipment Status'],inplace=True)#ahora, dropeo todos los np.nan, asi me quedo solo con los shipments recibidos y de mi sitio

#Esto... no me acuerdo porq pero messirve. Creo una DF solo de la info que me importa con las columnas q me importan
cols=list(df.columns)
IP_RETURN=df[[cols[0]]+[cols[7]]+[cols[8]]]#0= Return Shipment Number, 7=Creation Date ,  8 = Shipped By

#Una vez que tengo esto, reviso en el SFR si tengo los PL, Shipment confirmation y Return shipment

for index_IP, row_IP in IP_SHIPMENT.iteritems(): #Checkeo todos los Shipment number que obtuve del IP SHIPMENT.xlsx
    for index,row in SFR.iteritems(): #checkeo cada columnma y si estoy en el nivel 06.01.04
        if SFR.loc[index,'Ref Model ID']=='06.01.04':
            if IP_SHIPMENT.loc[index_IP,['Shipment Number']] in SFR.loc[index,'Document Name']: #Si el numero de envio esta en la string 
                #TODO
                if 'Packing List' in SFR.loc[index,'Document Name']:
                    SFR.loc[index,'Is the document present in the file structure? (Y/N)']='Y'
                    SFR.loc[index,'Is the document present in the file structure? (Y/N)']='Y'
                #Revisar todo esto. estoy laburando en una df pero despues tengo q modificar el excel. me parece q es al pedo esto y tenog q contrar una manera
                # de directamente modificar el excel. tal vez el index me sirve para ir directamente a la row que necesito modificar sin tener q ir
                # checkeando todo?







# #Grabo el nuevo excel con otro nombre
wb.save(Nombre_de_archivo)
