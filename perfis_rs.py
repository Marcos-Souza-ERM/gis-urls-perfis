
# -*- coding: utf-8 -*-
"""
Created on Mon May 15 10:29:50 2023

@author: samuel.matos
"""
import pandas as pd, sys
import os, pyshorteners
import math, webbrowser
from time import sleep, time
from datetime import datetime

start = time()
folder = r'C:\Users\samuel.matos\OneDrive - ERM\Samuel\Braskem BA\Perfis\Braskem RS' +'\\'
          

#type_tiny = pyshorteners.Shortener()
#type_tiny.tinyurl.short(long_url)

#IMPORTANTE OBSERVAR A PARTE DO NOME DO SITE



site = 'RS'
if site == 'BA':
    gis = pd.read_excel(folder +'BRASKEM_BA_PerfisConstrutivos_Controle.xlsx')
elif site == 'AL':
    gis = pd.read_excel(folder +'BRASKEM_AL_PerfisConstrutivos_Controle.xlsx')
elif site == 'Tupy':
    #nome do arquivo controle
    arquivo = 'TupyMaua_PerfisConstrutivos_Controle.xlsx'
    gis = pd.read_excel(folder +arquivo)
elif site == 'RS':
    arquivo = 'BRK-RS-Controle perfis das locations.xlsx'
    gis_sites = pd.ExcelFile(folder +arquivo)
    gis = pd.DataFrame()
    gis = pd.concat([ pd.read_excel(gis_sites, tab) for tab in gis_sites.sheet_names], ignore_index=True) 

# colunas da tabela de perfis enviada pelo gis (disponível no sharepoint)
gis_location_column = 'ID do Ponto'
site_column = 'Site'
subarea_column = 'Subárea'

['GlobalID *', 'OBJECTID', 'SYS_LOC_CODE', 'File_Name', 'Site',
       'Perfil_Disponivel', 'LinkPerfilConstrutivo']

links = pd.read_excel(folder + f'perfil {site}.xlsx')

#colunas da tabela (vem do BI)
['Name', 'Extension', 'Date accessed', 'Date modified', 'Date created',
       'Folder Path']

links['Name'] = links['Name'].str.strip()

links['Links_File_Name'] = links.Name.str[:-4]

links['Links_File_Name']         = links['Links_File_Name'].str.replace('|','I')
gis[gis_location_column]   = gis[gis_location_column].str.replace('|','I').str.strip()

#TODO

gis[site_column] = gis[site_column].str.strip()

#IMPORTANTE VERIFICAR O NOME DO SITE A PARTIR DO NOME DA PASTA
if site =='RS':
    gis.loc[gis[subarea_column].eq('-'), subarea_column] = ''
    links[site_column] = links['Folder Path'].str.split('Boring Logs - WebGIS').str[-1].str.strip('/')
    links[subarea_column] = ''
    links.loc[links[site_column].str.contains('/'), subarea_column] = links[site_column].str.split('/').str[-1]
    links[site_column] = links[site_column].str.strip()
    links.loc[links[site_column].str.contains('/'), site_column] = links[site_column].str.split('/').str[0]
    links['key'] = links[site_column] + links[subarea_column] + links['Links_File_Name']
    gis['key'] = gis[site_column]  + gis[subarea_column] + gis[gis_location_column]

else:
    links['site'] = links['Folder Path'].str.split('/').str[-2]
    links['key'] = links[site_column] + links['Links_File_Name']
    gis['key'] = gis[site_column]  + gis[gis_location_column]

not_in_sharepoint = gis[~gis.key.isin(links.key.tolist())]
not_in_gis_file   = links[~links.key.isin(gis.key.tolist())]

if any(not(df.empty) for df in [not_in_sharepoint, not_in_gis_file]):
    with pd.ExcelWriter('QC - Encurtamento de links.xlsx') as f:
        not_in_sharepoint.to_excel(f, 'Pontos excel n encontrados noSP', index=False)
        not_in_gis_file.to_excel(f,   'Pontos noSP n encontrados no ex', index=False)
    sleep(1)
    webbrowser.open('QC - Encurtamento de links.xlsx')
    
gis = pd.merge(gis, links[['Name', 'Folder Path','Links_File_Name','key']], 'left', 'key')



gis['link antigo'] = gis['Folder Path'] + gis['Name']
gis['link'] = None
go = True
#sys.exit(0)

if gis[~gis['link antigo'].isnull()].empty:
    print('nenhuma correspondência de entre as duas tabelas, não possível fazer o merge')
    sys.exit(0)
   
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


