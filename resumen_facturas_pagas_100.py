

# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 15:05:35 2023

@author: jcgarciam

El siguiente código tiene como fin enviar generar los archivos para la notificación de la DINA
de facturas 100 % pagas.
"""

import pandas as pd
import glob
import numpy as np
from datetime import datetime
import win32com.client as win32
import tkinter as tk



now = datetime.now()
# Rutas de las que se alimenta el código
path_int1 = r'\\DC1PVFNAS1\Autos\BusinessIntelligence\19-Soat-Salud-Arl\4-TRANSVERSAL\SISCO\SISCO\General\Salidas\SALUD'
path_int2 = r'\\DC1PVFNAS1\Autos\BusinessIntelligence\19-Soat-Salud-Arl\4-TRANSVERSAL\SISCO\SISCO\General\Salidas\ARL'
path_int3 = r'\\dc1pcadfrs1\Reportes_Activa'

# Ruta donde se guardan la salida de los archivos
path_salida = r'\\DC1PVFNAS1\Autos\BusinessIntelligence\19-Soat-Salud-Arl\4-TRANSVERSAL\SISCO\SISCO\Notificaciones DIAN\Output'

#%%
# La siguiente función nos permite generar una interfaz para ingresar manualmente la fecha de 
# los archivos que se cargaron, este fecha se debe ingresar con el formato ddmmaaaa
def ObtenerFecha():
    global fecha_archivos
    fecha_archivos = entrada.get()
    print('Fecha para la extracción de archivos:', fecha_archivos)
    ventana.destroy()
    
ventana = tk.Tk()
ventana.title('Fecha de los archivos cargados')
ventana.geometry("350x90")

etiqueta = tk.Label(ventana, text = 'Ingrese la fecha (ddmmaaaa),\n ejemplo 29022024:')
etiqueta.grid(row = 0, column = 0, padx = 10, pady = 5)

entrada = tk.Entry(ventana)
entrada.grid(row = 0, column = 1, padx = 10, pady = 5)

boton = tk.Button(ventana, text = 'Fecha', command = ObtenerFecha)
boton.grid(row = 1, padx = 10, pady = 5, columnspan = 2)

ventana.mainloop()


#%%
# EXTRACCIÓN
print('')
dic = {}

# Se cargan los archivos de acuses
for i in glob.glob(path_int3 + '/Acuses_*' + fecha_archivos + '*'):
    print('Leyendo el archivo: ', i[len(path_int3) + 1::])
    df = pd.read_csv(i, sep = ',', usecols = ['CUFE'])
    dic[i] = df
    print('Archivo ', i[len(path_int3) + 1::], ' leído\n')
    
acuses = pd.concat(dic).reset_index(drop = True)
#%%
dic = {}
# Se cargan los archivos de Recibidas
for i in glob.glob(path_int3 + '/Recibidas_*' + fecha_archivos + '*'):
    print('Leyendo el archivo: ', i[len(path_int3) + 1::])
    df = pd.read_csv(i, sep = ',')
    df['Origen'] = i[len(path_int3) + 1::]
    dic[i] = df
    print('Archivo ', i[len(path_int3) + 1::], ' leído\n')
    
recibidas = pd.concat(dic).reset_index(drop = True)


#%%
# Se cargan el archivo Maestro Salud que es SISCO Salud garupado por factura
columnas = ['Fecha_Radicacion','Valor_Neto','Total Valor Pagado','NIT','Numero_Factura','Regimen','Valor_Iva']
print('Cargando Maestro Salud')
Maestro_salud = pd.read_csv(path_int1 + '/Maestro_Salud.csv', sep = '|', usecols = columnas, encoding = 'ANSI')
print('Maestro Saludo cargado \n')
# Se cargan el archivo Maestro ARL que es SISCO ARL garupado por factura
print('Cargando Maestro ARL')
Maestro_arl = pd.read_csv(path_int2 + '/Maestro_ARL.csv', sep = '|', usecols = columnas, encoding = 'ANSI')
print('Maestro ARLL cargado \n')

sisco_agrupado = pd.concat([Maestro_salud,Maestro_arl]).reset_index(drop = True)
#%%

acuses['Acuses'] = 'Si'
acuses = acuses.rename(columns = {'CUFE':'Cufe'})
acuses['Cufe'] = acuses['Cufe'].astype(str)
recibidas['Cufe'] = recibidas['Cufe'].astype(str)

# Se cruza recibidas con acuses y solo se deja lo que cruce
recibidas2 = recibidas.merge(acuses, how = 'inner', on = 'Cufe')
#%%
sisco_agrupado['Valor_Neto'] = sisco_agrupado['Valor_Neto'].astype(str).str.replace(',','.').astype(float)
sisco_agrupado['Total Valor Pagado'] = sisco_agrupado['Total Valor Pagado'].astype(str).str.replace(',','.').astype(float)
sisco_agrupado['Valor_Iva'] = sisco_agrupado['Valor_Iva'].astype(str).str.replace(',','.').astype(float)

# Se evalúa la diferencia porcentual de lo cobrado versus lo pagado
sisco_agrupado['Porc Diferencias'] = (sisco_agrupado['Valor_Neto'] - sisco_agrupado['Total Valor Pagado'])/sisco_agrupado['Valor_Neto']
# Se evalúa la diferencia porcentual de lo cobrado versus lo pagado pero junto con el Valor Iva
# ya que hay facturas que se les debe agregar el Iva para que coincida el Pago versus lo cobrado
sisco_agrupado['Porc Diferencias2'] = (sisco_agrupado['Valor_Neto'] + sisco_agrupado['Valor_Iva'] - sisco_agrupado['Total Valor Pagado'])/(sisco_agrupado['Valor_Neto'] + sisco_agrupado['Valor_Iva'])

#%%

sisco_agrupado2 = sisco_agrupado.copy()
sisco_agrupado2['Fecha_Radicacion'] =  pd.to_datetime(sisco_agrupado2['Fecha_Radicacion'], format = '%Y/%m/%d')
sisco_agrupado2 = sisco_agrupado2[sisco_agrupado2['Fecha_Radicacion'].dt.year >= 2023]
sisco_agrupado2 = sisco_agrupado2[sisco_agrupado2['Valor_Neto'].abs() > 2]
#%%
sisco_agrupado2['Facturas Pagas Total'] = np.nan

sisco_agrupado2['Facturas Pagas Total'] = np.where(((sisco_agrupado2['Porc Diferencias'].abs() <= 0.01) | (sisco_agrupado2['Porc Diferencias2'].abs() <= 0.01)), 'Si', 'No')


#%%
def CambioFormato(df, a = 'a'):
    df[a] = df[a].astype(str)
    df[a] = np.where(df[a].str[-2::] == '.0', df[a].str[0:-2], df[a])
    df.loc[(df[a].str.contains('nan') == True),a] = np.nan

    return df[a]

sisco_agrupado3 = sisco_agrupado2[sisco_agrupado2['Facturas Pagas Total'] == 'Si']

sisco_agrupado3 = sisco_agrupado3[['NIT','Numero_Factura','Facturas Pagas Total','Regimen','Total Valor Pagado']]

sisco_agrupado3['NIT'] = CambioFormato(sisco_agrupado3, a = 'NIT')
sisco_agrupado3['Numero_Factura'] = CambioFormato(sisco_agrupado3, a = 'Numero_Factura')

sisco_agrupado3 = sisco_agrupado3[sisco_agrupado3['NIT'].isnull() == False]
sisco_agrupado3 = sisco_agrupado3[sisco_agrupado3['Numero_Factura'].isnull() == False]

sisco_agrupado3 = sisco_agrupado3.rename(columns = {'NIT':'ID Proveedor','Numero_Factura':'Número Documento'})

recibidas2['ID Proveedor'] = CambioFormato(recibidas2, a = 'ID Proveedor')
recibidas2['Número Documento'] = CambioFormato(recibidas2, a = 'Número Documento')

#%%
recibidas3 = recibidas2.merge(sisco_agrupado3, how = 'inner', on = ['ID Proveedor','Número Documento'], validate  = 'many_to_one')
recibidas3 = recibidas3.rename(columns = {'Número Documento':'Numero Documento'})
recibidas3['ID Cliente'] = CambioFormato(recibidas3, a = 'ID Cliente')
recibidas3['ID Cliente'] = recibidas3['ID Cliente'].str.strip('\ufeff').str.strip('"')
recibidas3['Comentario'] = np.nan
recibidas3['Estado documento'] = 'Aceptado'
#%%

for i in list(recibidas3['Origen'].unique()):
    print('\nGuardando archivo para ',i)
    df = recibidas3[recibidas3['Origen'] == i].reset_index(drop = True)
    print('\n    La base',i,'tiene', str(len(df)),'registros')
    
    cant = 2448
    a = len(df) % cant
    b = int((len(df) - a) / cant)
    if a > 0:
        b = b + 1
    print('    Se van a guardar',b,'archivos para',i,'\n')

    k = 0
    cant2 = cant
    
    for j in range(b):
            
        df2  = df.loc[k:cant2].copy()
        print('    Archivo',j,':','empieza en', k, 'y termina en',str(df2.index.max()) + '.','Tamaño:', df2.shape[0])
        k = cant2 + 1
        cant2 = cant + k

        df2 = df2.reset_index(drop = True)
        df2 = df2.reset_index()
        df2['index'] = df2['index'] + 1
        df2 = df2.rename(columns = {'index':'identificador'})
        df2 = df2[['identificador','ID Cliente','ID Proveedor','Tipo Documento','Numero Documento','Cufe','Comentario','Estado documento','Origen']]
        df2 = df2.drop(columns = ['Origen'])     
        df2.to_csv(path_salida + '/Salida ' + i[0:-4] + '_' + str(j) + '.csv', index = False, sep = ';', encoding = 'ANSI')

#%%

def correo(destinatarios, copia, asunto, cuerpo, adjunto = []):
    
    outlook_app = win32.Dispatch('Outlook.Application')

    mail_item = outlook_app.CreateItem(0)
    
    for i in adjunto:
        mail_item.Attachments.Add(i)
            
    
    mail_item.subject = asunto
    mail_item.Body = cuerpo
    mail_item.To = destinatarios
    mail_item.CC = copia
    mail_item.SentOnBehalfOfName = 'BusinessIntelligence@seguros.axacolpatria.co' 
    
    mail_item.Send()
    print('\nCorreo enviado, Notificaciones DIAN')
    
destinatarios = 'cristian.ochoa@axacolpatria.co;maria.mejiab@axacolpatria.co;ecbeltranc@axacolpatria.co;justo.gomez@axacolpatria.co;maria.luna@axacolpatria.co'
copia = 'BusinessIntelligence@seguros.axacolpatria.co;aurora.linares@axacolpatria.co'
asunto = 'MVP Generación datos de pago desde SISCO para notificación de eventos DIAN ' + fecha_archivos[0:2] + '-' + fecha_archivos[2:4] + '-' + fecha_archivos[4:8]
cuerpo = "Buen día,\n" + \
       "\n\n" + \
       'Se hace envío de la salida de los archivos para la notificación de facturas de MPP ' + \
       "y Vida cargados el día " + fecha_archivos[0:2] + '-' + fecha_archivos[2:4] + '-' + fecha_archivos[4:8] + '.'\
       "\n" + \
       "\n\n" + \
       "Cordialmente,\n Analítica de Siniestros \n"
adjunto = glob.glob(path_salida + '/Salida*' + fecha_archivos + '*')  

correo(destinatarios, copia, asunto, cuerpo, adjunto)        


        
print('\n')
print('Proceso finalizado')
print('Tiempo de duración: ', datetime.now() - now)
