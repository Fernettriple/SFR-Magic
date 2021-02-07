#!/usr/bin/env python3

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
import pandas as pd
import win32com.client as win32
import os
import numpy as np
import csv
from datetime import datetime as Noseporque


###
#TODO VER
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

# Guardo todo antes del procesador

wb.save(Nombre_de_archivo)

#TODO
#Procesado
#Abro el documento con Pandas

filename=os.getcwd()+'\\'+Nombre_de_archivo #consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa

SFR= pd.read_excel(filename, sheet_name='Site',header=0)
wb = openpyxl.load_workbook(Nombre_de_archivo)    
ws=wb['Site']
#Pongo las dates en formato datetime
def Str_to_date(str):
    try:
        return Noseporque.strptime(str,'%d-%b-%Y').date()
    except:
        return None
for index, row in SFR['Document date'].iteritems():
    SFR.loc[index,'Document date']=Str_to_date(SFR.loc[index,'Document date'])
    SFR.loc[index,'Expiration date']=Str_to_date(SFR.loc[index,'Expiration date'])


# if 'CONTACT REPORT.csv' in os.listdir('.'):
#     fname = os.getcwd()+"\\CONTACT REPORT.csv"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#     fname=fname.split('.')[0]+".xlsx"
#     wb.SaveAs(fname, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
#     wb.Close()                               #FileFormat = 56 is for .xls extension
#     excel.Application.Quit()
#     Contact_Report= pd.read_excel(fname,header=0)

#Escanear nombres en CV y SubI de los CVs para predecir el study team. Luego, pedir input del DOA y predecir q se tiene q pedir en base a lo que hay. 
# se podra usar un CTMS report?

#Buscar ultima FDA1572 y si es de hace dos años, preguntar si es la ultima.

#usando el reporte de visitas, checkear que estan todas las cfm, fup, svr

if 'VISIT REPORT.csv' in os.listdir('.') and 'VISIT REPORT.xlsx' not in os.listdir('.'):
    fname = os.getcwd()+"\\VISIT REPORT.csv"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    fname=fname.split('.')[0]+".xlsx"
    wb.SaveAs(fname, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    
class Sitio:
    def add_atribute(atribute, info_to_add):
        '''Esta funcion sirve para poder crear un atribute y asignarle un value llamado "info_to_add. Esto es especialmente util si estoy agregando cosas en un for loop'''
        if type(atribute)==str:
            atribute=atribute.replace(' ','_') #depuro los posibles caracteres q puedan joder al ponerle el atributo al objeto. TODO usar regex
            atribute=atribute.replace('/','_')
        if hasattr(Sitio,atribute):
            New_info=getattr(Sitio,atribute)
            New_info.append(info_to_add) #Siexiste el atributo, lo appendeo. Ya que siempre es una lista
        else:
            New_info=[info_to_add] #else, lo hago una lista.
        setattr(Sitio,atribute,New_info)
        
    class Investigador:
        pass

#Agarro del Visit report y agrego al Sitio cada una de las visitas, usando como nombre de atributo el tipo de visita, y el atributo es la fecha..
#El atributo es una lista, y si hay mas de una visita del mismo tipo, lo appendea
Visit_Report= pd.read_excel( os.getcwd()+"\\VISIT REPORT.xlsx",header=0)
for index_Visit_Report,row_Visit_Report in Visit_Report['Visit Type'].iteritems():    
    try:
        Visit_Report.loc[index_Visit_Report,'Visit End']=Visit_Report.loc[index_Visit_Report,'Visit End'].strftime('%d-%b-%Y')
        Sitio.add_atribute(Visit_Report['Visit Type'][index_Visit_Report], Visit_Report['Visit End'][index_Visit_Report])
    except:
        pass
#Ahora parseo por el DF del excel
def add_to_excel(Row_num,Present_in_eTMF,Comments,Action_needed,*Action):
    '''Esta Funcion sirve para agregar los comentarios al Excel. '''
    wb = openpyxl.load_workbook(Nombre_de_archivo)    
    ws=wb['Site']
    Row_num+=2 #para la df el primer index es 0, pero el excel arranca en 2
    if Present_in_eTMF=='N':
        Row_num=ws.max_row+1 #Si no esta presente, mando el comentario al fondo
    ws.cell(Row_num,11).value=Present_in_eTMF
    ws.cell(Row_num,12).value=Comments
    ws.cell(Row_num,13).value=Action_needed
    if Action_needed=='Y':
        ws.cell(Row_num,14).value=Action[0]
    wb.save(Nombre_de_archivo)

#Esto es para 05.04.03
def add_visit_from_report(Ref_ID, Generic_Variable_in_the_loop):
    '''Esta funcion checkea todas las visitas del reporte de visitas y se fija si estan en el archivo de SFR. Si estan, escribe en comments 
    diciendo que es la carta, si no esta agrega al final una linea con la info de que es lo que falta'''
    Letter_Types=['Confirmation Letter','Follow-up Letter', 'Monitoring Report']
    for index, row in SFR['Ref Model ID'].iteritems():
        if SFR.loc[index,'Document date']==None:
            continue
        if (SFR.loc[index,'Ref Model ID']==Ref_ID and
            SFR.loc[index,'Document date']== Str_to_date(Generic_Variable_in_the_loop)):
            if SFR.loc[index,'Ref Model Subtype'] not in Letter_Types:
                add_to_excel(index,'Y',f"Duplicated {(SFR.loc[index,'Ref Model Subtype'])} from {Generic_Variable_in_the_loop} visit",'Y','Errase Duplicated')
                continue #Si tengo un duplicado, no va a estar en letter types xq ya fue popeado. 
            else:
                Letter_Types.remove(SFR.loc[index,'Ref Model Subtype'])
                SFR.loc[index,'Ref Model ID']=np.nan
                add_to_excel(index,'Y',f"{(SFR.loc[index,'Ref Model Subtype'])} from {Generic_Variable_in_the_loop} visit",'N')
    if Letter_Types!=[]:
        add_to_excel(0,'N',f'{Letter_Types} missing from {Generic_Variable_in_the_loop} visit','Y','Collect from Site') #el primer argumento no importa en este caso, ya que se va a a setear igual al fondo

def check_and_add(code, atribute):
    '''Esta Funcion agarra un Ref ID y se fija si en el objeto Sitio tengo un tipo de visita que corresponda a ese ID. Si esta, ejecuta add_visit_from_report'''
    if hasattr(Sitio,atribute):
        for Visit_Report in getattr(Sitio,atribute):
            add_visit_from_report(code, Visit_Report)

#TODO Agregar todo los tipos de visita

check_and_add('05.01.04','Site_Visit_Selection')
check_and_add('05.03.01','Site_Visit_Initiation')            
check_and_add('05.04.03','Site_Visit_Interim')      
check_and_add('05.04.08','Telephone_Closeout' )


#los temp/CALIBRATION logs lo puedo checkear desde los reporte de IP vs IP RETURNED
#si es local o central tmb lo puedo sacar del log 





# ##IP SHIPMENTS/PL AND RETURNS


# #IP Shipment confirmation
# filename=os.getcwd()+"\\IP SHIPMENT.xlsx" #consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa
# df= pd.read_excel(filename, sheet_name='Sheet',header=2) #Importo el excel en pandas. Header=2 porq ahi estan los encabezados
# for index,row in df['Ship to Site Number'].iteritems(): #checkeo cada row de la columna "ship to site numbers"
#     if row!=Numero_de_sitio:
#         df.loc[index,'Ship to Site Number']=np.nan #Si la row no es del sitio q me interesa, le mando np.nan
#     if df.loc[index,'Shipment Status']!='Received': 
#         df.loc[index,'Shipment Status']=np.nan #Si el shipment no fue recibido, le mando np.nan
       
# df.dropna(subset=['Ship to Site Number','Shipment Status'],inplace=True) #ahora, dropeo todos los np.nan, asi me quedo solo con los shipments recibidos y de mi sitio

# #Esto... no me acuerdo porq pero messirve. Creo una DF solo de la info que me importa con las columnas q me importan
# cols=list(df.columns) 
# IP_SHIPMENT=df[[cols[0]]+[cols[8]]+[cols[9]]] #0= Shipment Number, 8 = Shipped Date, 9= Received Date

# #IP Return
# filename=os.getcwd()+"\\IP RETURN.xlsx"#consigo la direccion del archivo. con el os.getcwd() obtengo la dir del directorio donde esta el programa
# df= pd.read_excel(filename, sheet_name='Sheet',header=2)#Importo el excel en pandas. Header=2 porq ahi estan los encabezados
# for index,row in df['Ship from Site Number'].iteritems(): #checkeo cada row de la columna "ship to site numbers"
#     if row!=Numero_de_sitio:
#         df.loc[index,'Ship from Site Number']=np.nan #Si la row no es del sitio q me interesa, le mando np.nan
#     if df.loc[index,'Return Shipment Status']!='Received':
#         df.loc[index,'Return Shipment Status']=np.nan #Si el shipment no fue recibido, le mando np.nan
       
# df.dropna(subset=['Ship from Site Number','Return Shipment Status'],inplace=True)#ahora, dropeo todos los np.nan, asi me quedo solo con los shipments recibidos y de mi sitio

# #Esto... no me acuerdo porq pero messirve. Creo una DF solo de la info que me importa con las columnas q me importan
# cols=list(df.columns)
# IP_RETURN=df[[cols[0]]+[cols[7]]+[cols[8]]]#0= Return Shipment Number, 7=Creation Date ,  8 = Shipped By

# #Una vez que tengo esto, reviso en el SFR si tengo los PL, Shipment confirmation y Return shipment

# for index_IP, row_IP in IP_SHIPMENT.iteritems(): #Checkeo todos los Shipment number que obtuve del IP SHIPMENT.xlsx
#     for index,row in SFR.iteritems(): #checkeo cada columnma y si estoy en el nivel 06.01.04
#         if SFR.loc[index,'Ref Model ID']=='06.01.04':
#             if IP_SHIPMENT.loc[index_IP,['Shipment Number']] in SFR.loc[index,'Document Name']: #Si el numero de envio esta en la string 
#                 #TODO
#                 if 'Packing List' in SFR.loc[index,'Document Name']:
#                     SFR.loc[index,'Is the document present in the file structure? (Y/N)']='Y'
#                     SFR.loc[index,'Is the document present in the file structure? (Y/N)']='Y'
#                 #Revisar todo esto. estoy laburando en una df pero despues tengo q modificar el excel. me parece q es al pedo esto y tenog q contrar una manera
#                 # de directamente modificar el excel. tal vez el index me sirve para ir directamente a la row que necesito modificar sin tener q ir
#                 # checkeando todo?







# #Grabo el nuevo excel con otro nombre
