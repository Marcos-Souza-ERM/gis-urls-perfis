# -*- coding: utf-8 -*-
"""
Created on Mon Feb  6 15:18:19 2023

@author: samuel.matos
"""

import pandas as pd
import os
pa_yael = r'C:\Users\samuel.matos\Documents\OneDrive - ERM\Python\Braskem BA\Plano de Amostragem Braskem BA - Yael.xlsx'
pa_taise = r'C:\Users\samuel.matos\Documents\OneDrive - ERM\Python\Braskem BA\Plano de Amostragem Braskem BA - Taise.xlsx'
df_y = pd.read_excel(pa_yael)
df_t = pd.read_excel(pa_taise)

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
with pd.ExcelWriter("Planos de amostragem.xlsx") as arq:
    df_y.to_excel(arq, "PA Yael", index = False)
    df_t.to_excel(arq, "PA Taise", index = False)


