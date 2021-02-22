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
import re
pd.options.mode.chained_assignment = None  # default='warn'

###
#TODO VER
#MENSAJE DE BIENVENIDA, EXPLICAR QUE LOS NOMBRES DE LOS REPORTES TIENEN QUE ESTAR EN MAYUSCULAS Y ESO

#Esto me permite convertir la wea a .xlsx. lo saque de StackOverflow (NO CONSERVA H-LINKS)
#Hago lo mismo para los dos reportes de IP
if 'IP RETURN.xlsx' not in os.listdir('.'):
    if 'IP RETURN.xls' in os.listdir('.'):
        fname = os.getcwd()+"\\IP RETURN.xls"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

if 'IP SHIPMENT.xlsx' not in os.listdir('.'):
    if 'IP SHIPMENT.xls' in os.listdir('.'):
        fname = os.getcwd()+"\\IP SHIPMENT.xls"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()


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

#Creo este objeto para poder almacenar la info
class Sitio:
    def add_atribute(atribute, info_to_add):
        '''Esta funcion sirve para poder crear un atribute y asignarle un value llamado "info_to_add". Esto es especialmente util si estoy agregando cosas en un for loop'''
        if type(atribute)==str:
            atribute=atribute.replace(' ','_') #depuro los posibles caracteres q puedan joder al ponerle el atributo al objeto. TODO usar regex
            atribute=atribute.replace('/','_')
        if hasattr(Sitio,atribute):
            New_info=getattr(Sitio,atribute)
            New_info.append(info_to_add) #Siexiste el atributo, lo appendeo. Ya que siempre es una lista
        else:
            New_info=[info_to_add] #else, lo hago una lista.
        setattr(Sitio,atribute,New_info)

    #Informacion del Sitio
    Site_Number=Numero_de_sitio
    Cerrado = False
    #IP Shipment information
    First_IP=''    
    IP_Recieved=[]
    IP_Returned=[]

#TODO FDFs y Data privacy
#usando el reporte de visitas, checkear que estan todas las cfm, fup, svr
if 'VISIT REPORT.csv' in os.listdir('.') and 'VISIT REPORT.xlsx' not in os.listdir('.'):
    fname = os.getcwd()+"\\VISIT REPORT.csv"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    fname=fname.split('.')[0]+".xlsx"
    wb.SaveAs(fname, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    

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

def add_to_excel(Row_num,Ref_model_ID,Present_in_eTMF,Comments,Action_needed,*Action):
    '''Esta Funcion sirve para agregar los comentarios al Excel. '''
    wb = openpyxl.load_workbook(Nombre_de_archivo)    
    ws=wb['Site']
    if Present_in_eTMF=='N':
        Row_num=ws.max_row+1 #Si no esta presente, mando el comentario al fondo
    else:
        Row_num +=2 #para la df el primer index es 0, pero el excel arranca en 2
    ws.cell(Row_num,6).value = Ref_model_ID
    ws.cell(Row_num,11).value = Present_in_eTMF
    ws.cell(Row_num,12).value = Comments
    ws.cell(Row_num,13).value = Action_needed
    if Action_needed=='Y':
        ws.cell(Row_num,14).value=Action[0]
    wb.save(Nombre_de_archivo)

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
                add_to_excel(index,'05.04.03','Y',f"Duplicated {(SFR.loc[index,'Ref Model Subtype'])} from {Generic_Variable_in_the_loop} visit",'Y','Errase Duplicated')
                continue #Si tengo un duplicado, no va a estar en letter types xq ya fue popeado. 
            else:
                Letter_Types.remove(SFR.loc[index,'Ref Model Subtype'])
                SFR.loc[index,'Ref Model ID']=np.nan
                add_to_excel(index,'05.04.03','Y',f"{(SFR.loc[index,'Ref Model Subtype'])} from {Generic_Variable_in_the_loop} visit",'N')
    if Letter_Types!=[]:
        add_to_excel(0,'05.04.03','N',f'{Letter_Types} missing from {Generic_Variable_in_the_loop} visit','Y','Collect from Site') #el primer argumento no importa en este caso, ya que se va a a setear igual al fondo

def check_and_add(code, atribute):
    '''Esta Funcion agarra un Ref ID y se fija si en el objeto Sitio tengo un tipo de visita que corresponda a ese ID. Si esta, ejecuta add_visit_from_report'''
    if hasattr(Sitio,atribute):
        for Visit_Report in getattr(Sitio,atribute):
            add_visit_from_report(code, Visit_Report)
        if atribute == 'Site_Visit_Closeout' or atribute == 'Telephone_Closeout':
            Sitio.Cerrado = True

#TODO Agregar todo los tipos de visita (booster por ejemplo)

check_and_add('05.01.04','Site_Visit_Selection')
check_and_add('05.03.01','Site_Visit_Initiation')            
check_and_add('05.04.03','Site_Visit_Interim')  
check_and_add('05.04.08','Site_Visit_Closeout' )
check_and_add('05.04.08','Telephone_Closeout' )


#Extraigo informacion de IP SHIPMENT y lo meto en el Sitio
IP_SHIPMENT= pd.read_excel('IP SHIPMENT.xlsx', sheet_name='Sheet',header=2)

#Reduzco a mi sitio y a los envios recibidos
IP_SHIPMENT=IP_SHIPMENT.loc[IP_SHIPMENT['Shipment Status']=='Received']
IP_SHIPMENT_Site=IP_SHIPMENT.loc[IP_SHIPMENT['Ship to Site Number']==int(Sitio.Site_Number)]

if IP_SHIPMENT_Site.empty: #Puede que no haya tenido IP el sitio
    Sitio.IP_Recieved=None
    Sitio.First_IP=None
else:
    IP_Shipping_Dates=pd.to_datetime(IP_SHIPMENT_Site['Received Date'])

#Guardo todos los envios y el primero, porque me sirve para los IP temperature logs    
    Sitio.IP_Recieved=list(IP_SHIPMENT_Site['Shipment Number'])
    Sitio.First_IP=min(IP_Shipping_Dates).strftime('%d-%b-%Y')

#Ahora busco en SFR si estan los IP shipments
SFR_test=SFR.loc[SFR['Ref Model ID']=='06.01.04']
for shipment in Sitio.IP_Recieved:
    Shipment_types=['Packing List','Confirmation','Acknowledgement']
    Bacon=SFR_test.loc[SFR_test['Document Name'].str.contains(str(shipment), flags=re.IGNORECASE,na=False)]
    for documents in Shipment_types:
        if Bacon.index[Bacon['Document Name'].str.contains(documents)].empty==False:
            spam=Bacon.index[Bacon['Document Name'].str.contains(documents)]
            Shipment_types.remove(documents)
            if documents=='Acknowledgement':
                add_to_excel(spam[0],'06.01.04','Y',f"Check if this file is a Packing List, Shipping confirmation or Shipping Request",'N')
            else:
                add_to_excel(spam[0],'06.01.04','Y',f"{documents} for {shipment} shipping",'N')
    if Shipment_types!=[]:
        add_to_excel(0,'06.01.04','N',f'{Shipment_types} missing from {shipment} visit','Y','Collect from Site') #el primer argumento no importa en este caso, ya que se va a a setear igual al fondo


#Extraigo informacion de IP RETURN y lo meto en el Sitio
IP_RETURN= pd.read_excel('IP RETURN.xlsx', sheet_name='Sheet',header=2)
IP_RETURN=IP_RETURN.loc[IP_RETURN['Return Shipment Status']=='Received']
IP_RETURN_Site=IP_RETURN.loc[IP_RETURN['Ship from Site Number']==int(Sitio.Site_Number)]

if IP_RETURN_Site.empty: #Puede que no haya devuelto IP el sitio
    Sitio.IP_Returned=None  
else:
    IP_Return_Dates=pd.to_datetime(IP_RETURN_Site['Date Received'])

#Guardo los IP return
    Sitio.IP_Returned=list(IP_RETURN_Site['Return Shipment Number'])

SFR_test=SFR.loc[SFR['Ref Model ID']=='06.01.10']
if Sitio.IP_Returned != None:
    for shipment in Sitio.IP_Returned:    
        Bacon=SFR_test.loc[SFR_test['Document Name'].str.contains(str(shipment), flags=re.IGNORECASE,na=False)]
        if Bacon.index[Bacon['Document Name'].str.contains(str(shipment), flags=re.I, regex=True)].empty==False:
            spam=Bacon.index[Bacon['Document Name'].str.contains(str(shipment), flags=re.I, regex=True)]
            add_to_excel(spam[0],'06.01.10','Y',f"IP Return documentation for {shipment} shipping",'N')
        else:
            add_to_excel(0,'06.01.10','N',f"Missing IP Return Documentation for {str(shipment)} shipping",'Y','Collect from site')
else:    
    add_to_excel(SFR.index[SFR['Ref Model ID'] == '06.01.10'][0],'06.01.10','Y',f"No IP was returned",'N')
#Usando el primer Ip shipment, defino desde cuando necesito los IP temp logs y calibration logs

add_to_excel(0,'06.04.01','N',f"Please check that the IP temperature logs are present from {Sitio.First_IP} to present.",'Y','Collect from site, if applicable')
add_to_excel(0,'06.04.03','N',f"Please check that the calibration logs are present from {Sitio.First_IP} to present.",'Y','Collect from site, if applicable')

#TODO Predecir CVs, Med Lics, y GCPs
#usar un reporte de CTMS para predecir el study team (PIs, SubIs).
if 'CONTACT REPORT.csv' in os.listdir('.') and 'CONTACT REPORT.xlsx' not in os.listdir('.'):
    fname = os.getcwd()+"\\CONTACT REPORT.csv"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    fname=fname.split('.')[0]+".xlsx"
    wb.SaveAs(fname, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    Contact_Report= pd.read_excel(fname,header=0)
else:
    Contact_Report= pd.read_excel(os.getcwd()+"\\CONTACT REPORT.xlsx",header=0)
SFR= pd.read_excel(filename, sheet_name='Site',header=0)
Contact_Report=Contact_Report[['Role','First Name', 'Last Name', 'Start Date','End Date']]

class Site_Staff:
    """
    This class holds the information of the staff members

    Atributes:
        name (str): Name of the person.
        last_name (str): Last name of the person.
        role (str): Role of the person.
        start_date(timestamp): Starting date for that person on the site.
        end_date(timestamp): End date for that person on the site. This might not be present if the site member is still active.
        
    """

    def __init__ (self, name=None, last_name=None, role=None, start_date=None,end_date=None):
        """
        Constructor of Site_staff class.

        Parameters:
            name (str): Name of the person.
            last_name (str): Last name of the person.
            role (str): Role of the person.
            start_date(timestamp): Starting date for that person on the site.
            end_date(timestamp): End date for that person on the site. This might not be present if the site member is still active.
        """
        self.name = name
        self.last_name = last_name
        self.role= role
        self.GCP = False
        self.EDC = False
        self.IATA = False
        self.License = False
        if self.role == 'Principal Investigator':
            self.GCP = True
            self.EDC = True
            self.IATA = True
            self.License = True
        elif self.role == 'Sub-Investigator':
            self.GCP = True
            self.License = True
        self.start_date = start_date
        if pd.isna(end_date):
            self.end_date = 'Present'
        else:
            self.end_date = end_date

Names= Contact_Report['First Name'].tolist()
Last_names=Contact_Report['Last Name'].tolist()
Roles=Contact_Report['Role'].tolist()
Starting_dates=Contact_Report['Start Date'].tolist()
Ending_dates=Contact_Report['End Date'].tolist()

Site_Members = []

for nombres,apellidos,funcion,inicio,fin in zip(Names,Last_names,Roles,Starting_dates,Ending_dates):
    Site_Members.append(Site_Staff(nombres,apellidos,funcion,inicio,fin))

Sitio.Site_members=Site_Members

#Una vez que tengo la informacion guardada la uso para que haga cosas

#Una vez que tengo la informacion guardada la uso para que haga cosas

import datetime
pd.options.mode.chained_assignment = None  # default='warn'

SFR= pd.read_excel(filename, sheet_name='Site',header=0)
SFR['Document date']=pd.to_datetime(SFR['Document date'])
SFR['Expiration date']=pd.to_datetime(SFR['Expiration date'])

#Armo una df con lo q me interesa bc reasons
SFR_trainings= SFR.loc[(SFR['Ref Model ID'] == '05.02.07') | (SFR['Ref Model ID'] == '05.03.03')]

#Planteo los posibles certificados, previamente definidos en la clase
Certificates = ['GCP', 'EDC', 'IATA', 'License']

#Parseo por todos los staff members
for staff_member in Site_Members:    
    #Reduzco la df a solo lo que tiene el apellido del staff member en el nombre del archivo o en la columna de "site personnel name" (esta ultima a veces esta vacia xq la mtadata es malisima)
    df = SFR_trainings.loc[(SFR_trainings['Site personnel name'].str.contains(staff_member.last_name,na=False, flags=re.I, regex=True)) | (SFR_trainings['Document Name'].str.contains(staff_member.last_name,na=False, flags=re.I, regex=True))]
  
    #Por cada atributo en Certificates...    
    for atribute in Certificates: 
               
        #Si el atributo es True
        if getattr(staff_member,atribute) == True:
            
            #Me fijo si hay algun archivo presente que tenga el nombre del certificado (atribute) en el subtype o en el nombre
            #Si no hay nada, agrego una columna al final pidiendo lo que falta
            #Para evitar codigo feo, defino una nueva df_cert para no estar typeando df.loc[(df['Ref Model Subtype'].str.contains(atribute)) | (df['Document Name'].str.contains(atribute))]
            #todo el tiempo
            df_cert = df.loc[(df['Ref Model Subtype'].str.contains(atribute, flags=re.I, regex=True)) | (df['Document Name'].str.contains(atribute, flags=re.I, regex=True))]                
            if atribute == 'GCP' or atribute == 'License':
                Ref_model= '05.02.07'
            else:
                Ref_model= '05.03.03'
                
            #Si no encuentro resultados, agregar al fondo
            #TODO ARREGLAR TODOS LOS TEMAS DE LAS FECHAS DIOS ES UN CANCER TENER Q HACER ESTO:
            if staff_member.end_date == 'Present':
                msg = 'Present'
            elif type(staff_member.end_date) == str:
                msg = datetime.datetime.strptime(staff_member.end_date,"%d-%b-%Y").date()
            if type(staff_member.start_date) == str:
                staff_member.start_date = datetime.datetime.strptime(staff_member.start_date,"%d-%b-%Y")

            if df_cert.empty:                
                add_to_excel(' ',Ref_model, 'N', f'{atribute} for {staff_member.last_name} covering from {staff_member.start_date.date()} to {msg}', 'Y', 'Collect from site')
                  
            #Si encontro archivos vamos a checkear la fecha y compararla con lo que se necesita
            else:
                #ordeno la DF por fecha creciente
                df_cert.sort_values(by='Document date', inplace=True)
                df_cert.reset_index(inplace=True)
                
                #evaluo todos los items en la df
                #Para ir checkeando necesito ir trackeando las fechas cubiertas. Para esto creo                
                Cert_date = staff_member.start_date
                
                #ahora parseo por toda la df en orden creciente                
                for index in df_cert.index:  
                 #Como algunas certificaciones no tienen exp date porque la metadata es un sida, lo arreglo aca
                    if atribute == 'GCP':
                        df_cert['Expiration date'][index] = df_cert['Document date'][index]+datetime.timedelta(days=1095)
                    elif atribute == 'EDC':
                        df_cert['Expiration date'][index] = df_cert['Document date'][index]+datetime.timedelta(days=42069)
                    elif atribute == 'IATA': 
                        df_cert['Expiration date'][index] = df_cert['Document date'][index]+datetime.timedelta(days=730)
                    elif atribute == 'License' and pd.isna(df_cert.loc[index,'Expiration date']):
                        df_cert['Expiration date'][index] = df_cert['Document date'][index]+datetime.timedelta(days=365)                 
                    
                    #Ahora extraigo el index correspondiente al row en la SFR original y agrego la info en los comments
                    comment = f"{atribute} certificate from {df_cert['Document date'][index].date()} to {df_cert['Expiration date'][index].date()}"
                    add_to_excel(df_cert['index'][index],Ref_model,'Y',comment , 'N')  

                    #si la diferencia de fecha entre la licencia/training y la fecha de inicio/licencia anterior es mayor a 0, significa q el training ocurrio antes de la fecha limite, ergo esta todo bien
                    #Pero si es menor a 0, significa q el certificado se expidio despues de la fecha limite.
                    #seteo unos 90 de gracia para que la dif este todo bien, pero si es mayor a esos 90 dias hago cosas
                    if type(Cert_date) == str:
                        Cert_date = datetime.datetime.strptime(Cert_date,'%d-%b-%Y')
                    if (df_cert['Document date'][index] - Cert_date) > datetime.timedelta(days=90):                             
                        add_to_excel(df_cert['index'][index],Ref_model,'N',f"Missing {atribute} certificate for {staff_member.last_name}, {staff_member.name} covering from {Cert_date.date()} to {df_cert['Document date'][index].date()} missing", 'Y', 'Collect from site')
                    Cert_date = df_cert['Expiration date'][index]          
                    
                #checkeo la dif entre cuando vence la ultima licencia y cuando se fue del sitio o presente
                if staff_member.end_date == 'Present':
                     if (datetime.datetime.today() - Cert_date) > datetime.timedelta(days=0):
                        add_to_excel(0,Ref_model, 'N', f'Missing {atribute} certificate for {staff_member.last_name}, {staff_member.name} from {Cert_date.date()} to {msg}.', 'Y', 'Collect from site, if applicable')
                else:            
                    if (staff_member.end_date - Cert_date) > datetime.timedelta(days=0):
                        add_to_excel(0,Ref_model, 'N', f'Missing {atribute} certificate for {staff_member.last_name}, {staff_member.name} from {Cert_date.date()} to {msg}.', 'Y', 'Collect from site, if applicable')
                

#TODO si es local o central tmb lo puedo sacar del log (COMO?? CUANDO TENGAS IDEAS PLASMALAS)

#TODO PAs y IBs. Usando la visita de iniciacion puedo predecir que PAs/IBs tendria que tener. Puedo usar lo mismo para los irb approvals.
import json
if 'study_data.json' not in os.listdir('.'):
    Study_data = {}
    while True:
        print('Introduzca lo que quiere agregar')
        print('1.- Protocol Amendment\n2.- Investigator Brochure')
        mode = input()
        if mode == '':
            break
        while True:
            
            print('Escriba por favor la informacion en el formato "Version-Fecha, siendo la fecha escriba en formato YYYYMMDD"')
            print('Por ejemplo: "08-20200518"')
            print('Si desea volver al otro menu, presione enter')
            data = input()
            if data == '':
                break
            version = data.split('-')[0]
            fecha = data.split('-')[1]
            if len(fecha) != 8:
                print(f'La fecha introducida {fecha} parece que no esta en el formato YYYYMMDD. Por favor, re introducela')
                fecha = input()
            if mode == 1:
                Study_data['PA V'+version] = fecha
            else:
                Study_data['IB V'+version] = fecha
            with open('study_data.json', 'w') as json_file:
                json.dump(Study_data, json_file)

            print()
            print('Si desea introducir otro amendment')
        print('Si desea terminar, presione enter')

#FDA5172 estoy dormido, si funciona funciona
SFR= pd.read_excel(filename, sheet_name='Site',header=0)
SFR['Document date']=pd.to_datetime(SFR['Document date'])

SFR_FDA = SFR.loc[(SFR['Ref Model ID'] == '05.02.08') & ~(SFR['Document date'].isna() )]
SFR_FDA.sort_values(by='Document date', inplace=True)
SFR_FDA.reset_index(inplace=True)
if (SFR_FDA['Document date'][0] - datetime.datetime.strptime(Sitio.Site_Visit_Initiation[0], '%d-%b-%Y')) > datetime.timedelta(days=365):
    add_to_excel(SFR_FDA['index'][0], '05.02.08', 'N', f"First FDA1572 is from {SFR_FDA['Document date'][0].date()} but site had its Site Visit Initiation in {Sitio.Site_Visit_Initiation[0]}. Please check",'N')
SFR_FDA=SFR_FDA.tail(1)
SFR_FDA.reset_index(inplace=True)
if Sitio.Cerrado == True:
    if hasattr(Sitio,'Site_Visit_Closeout'):
        if (datetime.datetime.strptime(Sitio.Site_Visit_Closeout[0], '%d-%b-%Y') - SFR_FDA['Document date'][0]) > datetime.timedelta(days=365):
            add_to_excel(SFR_FDA['index'][0], '05.02.08', 'N', f"Last FDA1572 is from {SFR_FDA['Document date'][0].date()} but site had its Site Visit Closeout in {Sitio.Site_Visit_Closeout[0]}. Please check",'N')
    else:
        if (datetime.datetime.strptime(Sitio.Telephone_Closeout[0], '%d-%b-%Y') - SFR_FDA['Document date'][0]) > datetime.timedelta(days=365):
            add_to_excel(SFR_FDA['index'][0], '05.02.08', 'N', f"Last FDA1572 is from {SFR_FDA['Document date'][0].date()} but site had its Telephone Closeout in {Sitio.Telephone_Closeout[0]}. Please check", "N")        
elif (datetime.datetime.today() - SFR_FDA['Document date'][0]) > datetime.timedelta(days=365):
    add_to_excel(SFR_FDA['index'][0], '05.02.08', 'N', f"Last FDA1572 is from {SFR_FDA['Document date'][0].date()} but site had its Site Visit Initiation in {Sitio.Site_Visit_Initiation[0]}. Please check", "N")

#TODO encontrar manera que se me habia ocurrido pero ahora no de saber si el sitio es local o central lab/irb. y pedir todolo necesario, incluyendo membership list 













# #Grabo el nuevo excel con otro nombre
