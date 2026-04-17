# -*- coding: utf-8 -*-
"""
Created on Mon May 15 10:29:50 2023

@author: samuel.matos
"""
import pandas as pd
import os, pyshorteners
import math, webbrowser
from time import sleep, time
from datetime import datetime

start = time()
folder = r'C:\Users\samuel.matos\OneDrive - ERM\Samuel\Braskem BA\Perfis' +'\\'
          
os.listdir(folder)
#type_tiny = pyshorteners.Shortener()
#type_tiny.tinyurl.short(long_url)

#IMPORTANTE OBSERVAR A PARTE DO NOME DO SITE



site = 'AL'
if site == 'BA':
    gis = pd.read_excel(folder +'BRASKEM_BA_PerfisConstrutivos_Controle.xlsx')
elif site == 'AL':
    gis = pd.read_excel(folder +'BRASKEM_AL_PerfisConstrutivos_Controle.xlsx')
elif site == 'Tupy':
    #nome do arquivo controle
    arquivo = 'TupyMaua_PerfisConstrutivos_Controle.xlsx'
    gis = pd.read_excel(folder +arquivo)
#colunas da tabela
['GlobalID *', 'OBJECTID', 'SYS_LOC_CODE', 'File_Name', 'Site',
       'Perfil_Disponivel', 'LinkPerfilConstrutivo']
if 'File_Name' not in gis.columns:
    gis['File_Name'] = gis['SYS_LOC_CODE']
links = pd.read_excel(folder + f'perfil {site}.xlsx')
#colunas da tabela
['Name', 'Extension', 'Date accessed', 'Date modified', 'Date created',
       'Folder Path']

links['Name'] = links['Name'].str.replace(' ','')
links['File_Name'] = links.Name.str[:-4]

links['File_Name'] = links['File_Name'].str.replace('|','I')
gis['File_Name']   = gis['File_Name'].str.replace('|','I')


#links['site'] = links['Folder Path'].str.split('/').str[-2]
#links['site'] = links.site.str.replace('PE3','PE-3').str.replace('PE2','PE-2').str.replace('PE1','PE-1').str.replace('PVC1','PVC').str.replace('UA-3','UA3').str.replace('UA-4','UA4')

#TODO
#IMPORTANTE VERIFICAR O NOME DO SITE A PARTIR DO NOME DA PASTA

links['site'] = links['Folder Path'].str.split('/').str[-2].str.replace(' ','').str.replace('-','')
gis['Site'] = gis['Site'].str.replace(' ','').str.replace('-','')
print('links: ',links['site'].unique())
print('gis: ', gis.Site.unique())

links['key'] = links['site'] + links['File_Name']
gis['key'] = gis['Site']  + gis['File_Name']
gis = pd.merge(gis, links[['Name', 'Folder Path','File_Name','key','site']], 'left', 'key')

gis[gis.File_Name_y.isnull()][['GlobalID', 'ObjectID', 'SYS_LOC_CODE', 'File_Name_x', 'Site']
                              ].to_excel(folder+'Pontos da planilha de controle não encontrados no sharepoint.xlsx', index=False)
webbrowser.open(folder+'Pontos da planilha de controle não encontrados no sharepoint.xlsx')
gis['link antigo'] = gis['Folder Path'] + gis['Name']
gis['link'] = None
go = True

if gis[~gis['link antigo'].isnull()].empty:
    print('nenhuma correspondência de entre as duas tabelas, não possível fazer o merge')
    oi
    
for idx, row in gis[(~gis['link antigo'].isnull()) & (gis['link'].isnull())].iterrows():
    type_tiny = pyshorteners.Shortener()
    if type(gis.loc[idx, 'link']) != str:
           try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
           except:
               sleep(5)
               try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
               except:
                   sleep(10)
                   gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
                   

gis.loc[~gis['link'].isnull(), 'link'] = '<a href="' + gis['link'] + '">Disponível</a>'

gis.to_excel(folder + f"Links encurtados {site +'-' +str(datetime.now().date())}.xlsx", index = False)
webbrowser.open(folder + f"Links encurtados {site +'-' +str(datetime.now().date())}.xlsx")
print(time()-start)


