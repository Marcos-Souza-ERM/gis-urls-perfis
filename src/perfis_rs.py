
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
from urllib.parse import urlsplit, urlunsplit, quote, urlencode, parse_qsl
from src.config import env_choice, env_path, env_seconds, env_text, mapped_file

def normalize_url(url: str) -> str:
    parts = urlsplit(url)
    path = quote(parts.path, safe="/%")
    query = urlencode(parse_qsl(parts.query, keep_blank_values=True), quote_via=quote)
    return urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))

start = time()
folder = env_path('PERFIS_RS_INPUT_DIR')
folder_outputs = env_path('PERFIS_RS_OUTPUT_DIR')
site = env_choice('PERFIS_RS_SITE', ('BA', 'AL', 'AC', 'Tupy', 'RS'))
gis_location_column = env_text('PERFIS_RS_LOCATION_COLUMN')
site_column = env_text('PERFIS_RS_SITE_COLUMN')
subarea_column = env_text('PERFIS_RS_SUBAREA_COLUMN', allow_empty=site != 'RS')
retry_wait_1 = env_seconds('LINK_RETRY_WAIT_1')
retry_wait_2 = env_seconds('LINK_RETRY_WAIT_2')
date = str(datetime.now().date())
qc_file = os.path.join(folder_outputs, env_text('PERFIS_RS_QC_FILE', site=site, date=date))
output_file = os.path.join(folder_outputs, env_text('PERFIS_RS_OUTPUT_FILE', site=site, date=date))

#type_tiny = pyshorteners.Shortener()
#type_tiny.tinyurl.short(long_url)

#IMPORTANTE OBSERVAR A PARTE DO NOME DO SITE


arquivo = os.path.join(folder, mapped_file('PERFIS_RS_CONTROL_FILES', site))
if site == 'RS':
    gis_sites = pd.ExcelFile(arquivo)
    gis = pd.DataFrame()
    gis = pd.concat([ pd.read_excel(gis_sites, tab) for tab in gis_sites.sheet_names], ignore_index=True) 
else:
    gis = pd.read_excel(arquivo)

# colunas da tabela de perfis enviada pelo gis (disponível no sharepoint)

['GlobalID *', 'OBJECTID', 'SYS_LOC_CODE', 'File_Name', 'Site',
       'Perfil_Disponivel', 'LinkPerfilConstrutivo']

links = pd.read_excel(os.path.join(folder, env_text('PERFIS_RS_LINKS_FILE', site=site)))
links = links.rename(columns={
    env_text('PERFIS_RS_LINK_NAME_COLUMN'): 'Name',
    env_text('PERFIS_RS_LINK_PATH_COLUMN'): 'Folder Path',
})

#colunas da tabela (vem do BI)
['Name', 'Extension', 'Date accessed', 'Date modified', 'Date created', 'Folder Path']

links['Name'] = links['Name'].str.strip()

links['Links_File_Name'] = links.Name.str[:-4]

links['Links_File_Name'] = links['Links_File_Name'].str.replace('|','I')
gis[gis_location_column] = gis[gis_location_column].str.replace('|','I').str.strip()

#TODO

gis[site_column] = gis[site_column].str.strip()

print("Valores da Coluna site do GIS:\n")
print(gis[site_column].unique())

#IMPORTANTE VERIFICAR O NOME DO SITE A PARTIR DO NOME DA PASTA

if site =='RS':
    gis.loc[gis[subarea_column].eq('-'), subarea_column] = ''
    links[site_column] = links['Folder Path'].str.split(env_text('PERFIS_RS_RS_PATH_MARKER'), regex=False).str[-1].str.strip('/')
    links[subarea_column] = ''
    links.loc[links[site_column].str.contains('/'), subarea_column] = links[site_column].str.split('/').str[-1]
    links[site_column] = links[site_column].str.strip()
    links.loc[links[site_column].str.contains('/'), site_column] = links[site_column].str.split('/').str[0]
    links['key'] = links[site_column] + links[subarea_column] + links['Links_File_Name']
    gis['key'] = gis[site_column]  + gis[subarea_column] + gis[gis_location_column]
elif site == 'BA':
    links['site'] = links['Folder Path'].str.split('/').str[-2]

    print("Valores da Coluna site do BI:\n")
    print(links['site'].unique())
    links['site'] = links.site.str.replace('PE3','PE-3').str.replace('PE2','PE-2').str.replace('PE1','PE-1').str.replace('PVC1','PVC').str.replace('UA-3','UA3').str.replace('UA-4','UA4')
    print(links['site'].unique())

    links['key'] = links['site'] + links['Links_File_Name']
    gis['key'] = gis[site_column]  + gis[gis_location_column]
elif site == 'AC': #AC
    links['site'] = links['Folder Path'].str.split(env_text('PERFIS_RS_AC_PATH_MARKER'), regex=False).str[-1].str.split('/').str[0]
    links['site'] = links.site.str.replace('01. REFMAT','REFMAT').str.replace('02. FASF','FASF')


    #links[subarea_column] = links['Folder Path'].str.split('/').str[-2]
    #links[subarea_column] = links.site.str.split(' .').str[1]
    #print("Valores da Subáreas do BI:\n")
    #print(links[subarea_column].unique())

    print("----------------------------------------------\n")

    print("Valores da Coluna site do BI:\n")
    print(links['site'].unique())

    print("----------------------------------------------\n")

    print("Valores da Coluna site do GIS:\n")
    print(gis[site_column].unique())


    links['key'] = links['site'] + links['Links_File_Name']
    gis['key'] = gis[site_column] + gis[gis_location_column]
else:
    links['site'] = links['Folder Path'].str.split('/').str[2]

    print("Valores da Coluna site do BI:\n")
    print(links['site'].unique())

    links['key'] = links['site'] + links['Links_File_Name']
    gis['key'] = gis[site_column]  + gis[gis_location_column]

# comparação entre as tabelas para verificar se existem pontos na tabela do GIS que não possuem correspondência na tabela de links (e vice-versa).
# Caso existam, um arquivo de QC será gerado listando os pontos que não possuem correspondência em cada tabela.

not_in_sharepoint = gis[~gis.key.isin(links.key.tolist())]
not_in_gis_file   = links[~links.key.isin(gis.key.tolist())]

if any(not(df.empty) for df in [not_in_sharepoint, not_in_gis_file]):
    with pd.ExcelWriter(qc_file) as f:
        not_in_sharepoint.to_excel(f, sheet_name='Pontos excel n encontrados noSP', index=False)
        not_in_gis_file.to_excel(f, sheet_name='Pontos noSP n encontrados no ex', index=False)
    sleep(1)
    webbrowser.open(qc_file)

print("Executando merge entre as tabelas...")

gis = pd.merge(gis, links[['Name', 'Folder Path','Links_File_Name','key']], 'left', 'key')

gis['link antigo'] = gis['Folder Path'] + gis['Name']
gis['link'] = None
go = True
#sys.exit(0)

if gis[~gis['link antigo'].isnull()].empty:
    print('nenhuma correspondência de entre as duas tabelas, não possível fazer o merge')
    sys.exit(0)

print("Iniciando processo de encurtamento de links...")

# o processo de encurtamento de links pode apresentar falhas quando muitos links são encurtados em um curto período de tempo, por isso,
# foi implementada uma lógica de tentativas para garantir que o processo seja concluído mesmo que ocorram falhas temporárias na conexão ou no serviço de encurtamento de links.


for idx, row in gis[(~gis['link antigo'].isnull()) & (gis['link'].isnull())].iterrows():

    if type(gis.loc[idx, 'link']) != str:
        type_tiny = pyshorteners.Shortener()
        url = normalize_url(gis.loc[idx, 'link antigo'])

        try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(url)
        except:
            sleep(retry_wait_1)
            try: gis.loc[idx, 'link'] = type_tiny.tinyurl.short(url)
            except:
                sleep(retry_wait_2)
                gis.loc[idx, 'link'] = type_tiny.tinyurl.short(url)


# transformação do link encurtado em um hyperlink clicável no excel

gis.loc[~gis['link'].isnull(), 'link'] = '<a href="' + gis['link'] + '">Disponível</a>'

gis.to_excel(output_file, index = False)
webbrowser.open(output_file)

print("Processo finalizado. Arquivo salvo em: " + output_file)
print(time()-start)
