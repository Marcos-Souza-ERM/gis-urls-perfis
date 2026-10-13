# -*- coding: utf-8 -*-
"""
Created on Fri Feb  3 18:02:42 2023

@author: samuel.matos
"""
import pandas as pd
import os
import webbrowser
import openpyxl as xl
from src.config import env_mapping, env_path, env_text

file = env_path('MAIN_PLAN_FILE')
folder = env_path('MAIN_EQUIS_DIR')
folder_checkl = env_path('MAIN_CHECKLIST_DIR')
plan_output_file = env_path('MAIN_PLAN_OUTPUT_FILE')
file_param_ausentes = env_path('MAIN_REPORT_OUTPUT_FILE')
sites = env_mapping('MAIN_EQUIS_FILES')
sites_check = env_mapping('MAIN_CHECKLIST_FILES')
site_rename = env_mapping('MAIN_SITE_RENAME')
plan_sheet_filter = env_text('MAIN_PLAN_SHEET_FILTER')
plan_columns = {
    env_text('MAIN_GROUP_COLUMN'): 'Grupos',
    env_text('MAIN_POINTS_COLUMN'): 'Poços',
    env_text('MAIN_PARAMETERS_COLUMN'): 'Parâmetros',
}
sample_sheet = env_text('MAIN_SAMPLE_SHEET')
smp_sheet = env_text('MAIN_SMP_SHEET')
test_sheet = env_text('MAIN_TEST_SHEET')
result_sheet = env_text('MAIN_RESULT_SHEET')
batch_sheet = env_text('MAIN_BATCH_SHEET')


fileobj = pd.ExcelFile(file)
param_array=[]
grupo_array=[]
ponto_array=[]
aba_array=[]
for aba in fileobj.sheet_names:
    if plan_sheet_filter in aba:
        df = pd.read_excel(fileobj, aba) 
        df = df.rename(columns = {'Parâmetros ':'Parâmetros'})
        df = df.rename(columns=plan_columns)
        for idx, row in df.iterrows():
            grupo = df.loc[idx, 'Grupos'].strip()
            #site_pa = df.loc[idx, 'Site'].strip()

            for ponto in df.loc[idx, 'Poços'].replace('\n','').replace(' ','').replace('"','').replace('PM-057/038PM-057/039','PM-057/038,PM-057/039').split(','):
                #Correção de nomes de parâmetros
                parametros = str(df.loc[idx, 'Parâmetros']).replace('\n','').replace('Benzo(g,h,i)perileno','Benzo(g.h.i)perileno').replace(
                    'Dibenzo(a,h)antraceno','Dibenzo(a.h)antraceno').replace('Arsênio, Dissolvido','Arsênio Dissolvido').replace('Cr+6, Cromo Total','Cr+6 , Cromo Total').replace(
                        'Tetracloretode carbono','Tetracloreto de carbono').replace('Tetracloreto decarbono','Tetracloreto de carbono').replace(
                            'Cloretode vinila','Cloreto de vinila').replace('1,2- Dicloroeteno(cis)','1,2-Dicloroeteno(cis)').replace(
                                'Hexaclooetano','Hexacloroetano').replace('CobreDissolvido','Cobre Dissolvido').replace('ManganêsTotal','Manganês Total').replace(
                                    'PrataDissolvido','Prata Dissolvido').replace('2,2-"Dicloropropano','2,2-Dicloropropano')
                            
                for i in range(0,10):
                    parametros = parametros.replace(str(i)+',',str(i)+'.')
                for param in parametros.split(','):
                    param_array.append(param.strip())
                    grupo_array.append(grupo.strip())
                    ponto_array.append(ponto.strip())
                    aba_array.append(aba)
            
            
data = pd.DataFrame(data = {'Ponto':ponto_array,
                            'Parâmetro':param_array,
                            'Grupos':grupo_array,
                            'Site':aba_array})   
data['Parâmetro'] = data['Parâmetro'].str.replace('.',',')   
data.loc[data['Parâmetro'].eq('nan'), 'Parâmetro'] = 'Parâmetro não informado'
#Correção de parâmetro
data.loc[data['Parâmetro'].eq('romometano'), 'Parâmetro'] = 'Bromometano'
data.loc[data['Parâmetro'].eq('TPH Alifático (C6- C8)'), 'Parâmetro'] = 'TPH Alifático (C6-C8)'
data.loc[data['Parâmetro'].eq('TPH GRO'), 'Parâmetro'] = 'TPH GRO (C6-C10)'
data.loc[data['Parâmetro'].eq('Cobalto "Total'), 'Parâmetro'] = 'Cobalto Total'



for i in site_rename.keys():
    data.loc[data.Site.eq(i), 'Site'] = site_rename[i]
#segunda campanha
input("""Lembre de considerar a campanha!
      ***
      """)
      
data = data[data.Site.isin(sites.keys())]
data.to_excel(plan_output_file, index = False)
"""
CHECKLIST
"""
data_checkl = pd.DataFrame()
for check_file in os.listdir(folder_checkl):
    checkl = pd.read_excel('\\'.join([folder_checkl, check_file]))
    
    for i in range(3):
        if list(set(checkl.iloc[:,i].fillna('')))[0] == '': checkl = checkl.drop(columns = checkl.columns[i])
    checkl.columns = checkl.loc[3]
    checkl = checkl.rename(columns = {checkl.columns[0]:'Poço',
                                      checkl.columns[1]:'Data'})
    checkl = checkl.loc[i:,['Poço','Data','Poço foi amostrado?','Status do Poço']].reset_index(drop = True)
    if not(len([i for i in ['Poço','Data','Poço foi amostrado?','Status do Poço'] if i in checkl.columns ])) ==4: oi #erro intencional
    checkl['Site'] = check_file
    data_checkl = pd.concat([data_checkl,checkl], ignore_index = True)
    for key in sites_check:
        data_checkl.loc[data_checkl.Site.eq(sites_check[key]),'Site'] = key
   
pa_extras           = pd.DataFrame()
pa_ausentes         = pd.DataFrame()
pontos_extras       = pd.DataFrame()
pontos_ausentes     = pd.DataFrame()
parametros_extras   = pd.DataFrame()
parametros_ausentes = pd.DataFrame()
compilado_result = pd.DataFrame()

for site in data.Site.unique():
    aba_sample = pd.read_excel(os.path.join(folder, sites[site]), sample_sheet)
    aba_smp = pd.read_excel(os.path.join(folder, sites[site]), smp_sheet)
    aba_test = pd.read_excel(os.path.join(folder, sites[site]), test_sheet)
    aba_result = pd.read_excel(os.path.join(folder, sites[site]), result_sheet)
    aba_batch = pd.read_excel(os.path.join(folder, sites[site]), batch_sheet)
    
    data_site = data[data['Site'].eq(site)]
    pontos_extras = pd.concat([pontos_extras,
                        aba_sample[~aba_sample['sys_loc_code'].isin(data_site.Ponto.tolist())]], ignore_index=True)
    pontos_ausentes = pd.concat([pontos_ausentes,data[~data['Ponto'].isin(aba_sample.sys_loc_code.tolist())]], ignore_index=True)
    parametros_extras = pd.concat([parametros_extras,aba_result[~aba_result['chemical_name'].isin(data_site.Parâmetro.tolist())]], ignore_index=True)
    parametros_ausentes = pd.concat([parametros_ausentes,data[~data['Parâmetro'].isin(aba_result.chemical_name.tolist())]], ignore_index=True)
    aba_result = pd.merge(aba_sample[['sys_sample_code','sys_loc_code']],aba_result,how= 'left',on='sys_sample_code').drop_duplicates()
    aba_result = aba_result[~aba_result['sys_sample_code'].eq('#Text(40)')]
    aba_result['Site'] = site
    compilado_result = pd.concat([compilado_result,aba_result], ignore_index = True)
    for ponto in data_site.Ponto.unique():
        samples = aba_result[aba_result['sys_loc_code'].eq(ponto)]['sys_sample_code'].unique()
        pa = data_site[data_site.Ponto.eq(ponto)]
        
        for sample in samples:
            pa_extras = pd.concat([pa_extras,
                                   aba_result[(aba_result['sys_sample_code'].eq(sample)) &
                                (~aba_result['chemical_name'].isin(pa.Parâmetro.tolist()))]], ignore_index=True)
            pa_ausentes_copy = pa[~pa.Parâmetro.isin(
                aba_result[aba_result['sys_sample_code'].eq(sample)].chemical_name.tolist())
                ]
            pa_ausentes_copy['sys_sample_code'] = sample
            pa_ausentes_copy['sample_date'] = aba_sample[aba_sample['sys_sample_code'].eq(sample)]['sample_date'].unique()[0]
            pa_ausentes_copy['Ponto no EDD?'] = 'Sim'
            pa_ausentes = pd.concat([pa_ausentes,
                                     pa_ausentes_copy], ignore_index=True)
        if len(samples) == 0:
            print(ponto)
            pa['Ponto no EDD?'] = 'Não'
            pa['Motivo'] = 'Sem justificativa'
            pa_ausentes = pd.concat([pa_ausentes, pa], ignore_index=True)

    if 'Site' in pontos_extras.columns:
        pontos_extras.loc[pontos_extras['Site'].isnull(), 'Site'] = site
        parametros_extras.loc[parametros_extras['Site'].isnull(),'Site'] = site
        pa_extras.loc[pa_extras['Site'].isnull(),'Site'] = site
    else:
        pontos_extras['Site'] = site
        parametros_extras['Site'] = site
        pa_extras['Site'] = site

nao_amostrado = data_checkl[data_checkl['Poço foi amostrado?'].str.lower().str.contains('não')]
for idx, row in nao_amostrado.iterrows():
    pa_ausentes.loc[(pa_ausentes['Site'].eq(row['Site'])) & 
                    (pa_ausentes['Ponto'].eq(row['Poço'])),'Motivo'] = row['Status do Poço']
amostrado = data_checkl[data_checkl['Poço foi amostrado?'].str.lower().str.contains('sim')]
amostrado['key'] = amostrado['Poço'] + '_' + amostrado['Site']
compilado_result['key'] = compilado_result['sys_loc_code'] + "_" + compilado_result['Site']
nao_amostrado_sem_just = amostrado[~amostrado.key.isin(compilado_result.key.tolist())]


with pd.ExcelWriter(file_param_ausentes) as arq:
    pd.DataFrame(data = {'Resumo':["""Pontos importantes: Ao em Análises ausentes Ponto no EDD? = Não, você conseguirá ver os pontos que não foram encontrados no EDD.
                                   Isto pode ocorrer devido ao nome do ponto estar escrito errado ou de fato o ponto não veio no EDD. Após ter utilizado o filtro acima, 
                                   você também irá ver se o motivo do ponto não ter sido amostrado, se houver "Sem justificativa" nesta coluna, significa que o laboratório
                                   ainda não justificou através dos checklist a não coleta deste ponto.""",
                                   'Ausentes = está no plano de amostragem e não está no EDD',
                                   'Extras = está no EDD mas não está no plano de amostragem']}).to_excel(arq, "Resumo", index = False)
    pa_extras.drop_duplicates().to_excel(arq, "Análises extras", index = False)
    pa_ausentes.drop_duplicates().to_excel(arq, "Análises ausentes", index = False)
    pontos_extras.drop_duplicates().to_excel(arq, "Pontos extras", index = False)
    pontos_ausentes.drop_duplicates().to_excel(arq, "Pontos ausentes", index = False)
    parametros_extras.drop_duplicates().to_excel(arq, "Nomes não encontrados no PA", index = False)
    parametros_ausentes.drop_duplicates().to_excel(arq, "Nomes não encontrados no edd", index = False)
    compilado_result[~compilado_result['sys_loc_code'].isin(data.Ponto.tolist())]['sys_loc_code'].drop_duplicates().to_excel(arq, "Pontos extras geral", index = False)
    data[~data['Ponto'].isin(compilado_result.sys_loc_code.tolist())]['Ponto'].drop_duplicates().to_excel(arq, "Pontos ausentes geral", index = False)
    compilado_result[~compilado_result['chemical_name'].isin(data.Parâmetro.tolist())]['chemical_name'].drop_duplicates().to_excel(arq, "Parâmetros extras geral", index = False)
    data[~data['Parâmetro'].isin(compilado_result.chemical_name.tolist())]['Parâmetro'].drop_duplicates().to_excel(arq, "Parâmetros ausentes geral", index = False)
    nao_amostrado_sem_just.drop_duplicates().to_excel(arq, "Não coletado sem justificativa", index = False)
    compilado_result.drop_duplicates().to_excel(arq, "Compilado EDDs", index = False)
    data.drop_duplicates().to_excel(arq, "Plano de amostragem", index = False)
    #['PVC_PA', 'Tecnor_PA', 'Q1_PA', 'PE-1_PA', 'PE-2_PA', 'PE-3_PA', 'UA-3_PA', 'UA-4_PA']
    data_checkl.to_excel(arq, "Check list", index = False)
    
wb = xl.load_workbook(file_param_ausentes)
wb['Resumo'].sheet_properties.tabColor = 'C40043'
wb['Análises ausentes'].sheet_properties.tabColor = 'C40043'
wb["Não coletado sem justificativa"].sheet_properties.tabColor = 'C40043'
wb.save(file_param_ausentes)


webbrowser.open(os.getcwd())
    
    
    
    
    



