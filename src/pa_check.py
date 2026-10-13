# -*- coding: utf-8 -*-
"""
Created on Mon Feb  6 15:18:19 2023

@author: samuel.matos
"""

import pandas as pd
import os
from src.config import env_path, env_text

pa_yael = env_path('PA_CHECK_YAEL_FILE')
pa_taise = env_path('PA_CHECK_TAISE_FILE')
output_file = env_path('PA_CHECK_OUTPUT_FILE')
columns = {
    env_text('PA_CHECK_SITE_COLUMN'): 'Site',
    env_text('PA_CHECK_POINT_COLUMN'): 'Ponto',
    env_text('PA_CHECK_PARAMETER_COLUMN'): 'Parâmetro',
}
df_y = pd.read_excel(pa_yael)
df_t = pd.read_excel(pa_taise)
df_y = df_y.rename(columns=columns)
df_t = df_t.rename(columns=columns)

df_t['key'] = df_t['Site'] + '-' + df_t['Ponto'] + '-' + df_t['Parâmetro'] 
df_y['key'] = df_y['Site'] + '-' + df_y['Ponto'] + '-' + df_y['Parâmetro'] 
df_t['Parâmetros ausentes'] = ''
df_y['Parâmetros ausentes'] = ''
df_t.loc[~df_t.key.isin(df_y['key'].tolist()),'Parâmetros ausentes'] = "Parâmetros solicitados no PA Taise que não foram solicitados no PA Yael"
df_y.loc[~df_y.key.isin(df_t['key'].tolist()),'Parâmetros ausentes'] = "Parâmetros solicitados no PA Yael que não foram solicitados no PA Taise"
df_t.loc[~df_t.Ponto.isin(df_y['Ponto'].tolist()),'Pontos ausentes'] = "Pontos solicitados no PA Taise que não foram solicitados no PA Yael"
df_y.loc[~df_y.Ponto.isin(df_t['Ponto'].tolist()),'Pontos ausentes'] = "Pontos solicitados no PA Yael que não foram solicitados no PA Taise"
df_t.loc[~df_t.Parâmetro.isin(df_y['Parâmetro'].tolist()),'Parâmetro ausentes geral'] = "Parâmetro solicitados no PA Taise que não foram solicitados no PA Yael em nenhum momento"
df_y.loc[~df_y.Parâmetro.isin(df_t['Parâmetro'].tolist()),'Parâmetro ausentes geral'] = "Parâmetro solicitados no PA Yael que não foram solicitados no PA Taise em nenhum momento"
with pd.ExcelWriter(output_file) as arq:
    df_y.to_excel(arq, "PA Yael", index = False)
    df_t.to_excel(arq, "PA Taise", index = False)
