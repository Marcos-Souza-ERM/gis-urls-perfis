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
from src.config import env_choice, env_path, env_seconds, env_text, mapped_file

start = time()
folder = env_path('PERFIS_INPUT_DIR')
folder_outputs = env_path('PERFIS_OUTPUT_DIR')
site = env_choice('PERFIS_SITE', ('BA', 'AL', 'Tupy'))
retry_wait_1 = env_seconds('LINK_RETRY_WAIT_1')
retry_wait_2 = env_seconds('LINK_RETRY_WAIT_2')
date = str(datetime.now().date())
qc_file = os.path.join(folder_outputs, env_text('PERFIS_QC_FILE', site=site, date=date))
output_file = os.path.join(folder_outputs, env_text('PERFIS_OUTPUT_FILE', site=site, date=date))
          
os.listdir(folder)
#type_tiny = pyshorteners.Shortener()
#type_tiny.tinyurl.short(long_url)

#IMPORTANTE OBSERVAR A PARTE DO NOME DO SITE

gis = pd.read_excel(os.path.join(folder, mapped_file('PERFIS_CONTROL_FILES', site)))
gis = gis.rename(columns={
    env_text('PERFIS_FILE_COLUMN'): 'File_Name',
    env_text('PERFIS_LOCATION_COLUMN'): 'SYS_LOC_CODE',
    env_text('PERFIS_SITE_COLUMN'): 'Site',
    env_text('PERFIS_GLOBAL_ID_COLUMN'): 'GlobalID',
    env_text('PERFIS_OBJECT_ID_COLUMN'): 'ObjectID',
})
#colunas da tabela
['GlobalID *', 'OBJECTID', 'SYS_LOC_CODE', 'File_Name', 'Site',
       'Perfil_Disponivel', 'LinkPerfilConstrutivo']
if 'File_Name' not in gis.columns:
    gis['File_Name'] = gis['SYS_LOC_CODE']
links = pd.read_excel(os.path.join(folder, env_text('PERFIS_LINKS_FILE', site=site)))
links = links.rename(columns={
    env_text('PERFIS_LINK_NAME_COLUMN'): 'Name',
    env_text('PERFIS_LINK_PATH_COLUMN'): 'Folder Path',
})
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
                              ].to_excel(qc_file, index=False)
webbrowser.open(qc_file)
gis['link antigo'] = gis['Folder Path'] + gis['Name']
gis['link'] = None
go = True

if gis[~gis['link antigo'].isnull()].empty:
    print('nenhuma correspondência de entre as duas tabelas, não possível fazer o merge')
    
for idx, row in gis[(~gis['link antigo'].isnull()) & (gis['link'].isnull())].iterrows():
    type_tiny = pyshorteners.Shortener()
    if type(gis.loc[idx, 'link']) != str:
           try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
           except:
               sleep(retry_wait_1)
               try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
               except:
                   sleep(retry_wait_2)
                   gis.loc[idx, 'link'] = type_tiny.tinyurl.short(gis.loc[idx,'link antigo'])
                   

gis.loc[~gis['link'].isnull(), 'link'] = '<a href="' + gis['link'] + '">Disponível</a>'

gis.to_excel(output_file, index = False)
webbrowser.open(output_file)
print(time()-start)
