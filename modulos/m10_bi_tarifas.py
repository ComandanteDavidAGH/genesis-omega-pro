import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import gspread
import re
import math
import io
import openpyxl

# --- 🧪 APARTADO DE BARREDORAS Y AUXILIARES GLOBALES (ESTRUCTURA PLANA) ---
def limpiar_encabezados(df):
    df.columns = [
        str(col).upper()
        .replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
        .replace('À','A').replace('È','E').replace('Ì','I').replace('Ò','O').replace('Ù','U')
        .strip()
        for col in df.columns
    ]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    if "" in df.columns: df = df.drop(columns=[""])
    return df
    
def estandarizar_base(df):
    renombres = {}
    for col in df.columns:
        col_u = str(col).upper().replace('\n', ' ').strip()
        if 'FACTURAR' in col_u:
            renombres[col] = 'COSTO_MAESTRO'
            break
            
    if 'COSTO_MAESTRO' not in renombres.values():
        for col in df.columns:
            col_u = str(col).upper().replace('\n', ' ').strip()
            if 'COSTO AVION ($/HA)' in col_u or col_u == 'COSTO_HA':
                renombres[col] = 'COSTO_MAESTRO'
                break
                
    finca_ok = False; fecha_ok = False; area_ok = False
    for col in df.columns:
        col_u = str(col).upper().replace('\n', ' ').strip()
        if not finca_ok and (col_u == 'FINCA' or col_u == 'PROPIEDAD'):
            renombres[col] = 'FINCA_MAESTRA'
            finca_ok = True
        elif not fecha_ok and col_u == 'FECHA':
            renombres[col] = 'FECHA_MAESTRA'
            fecha_ok = True
        elif not area_ok and ('FUMIG' in col_u or ('AREA' in col_u and 'BRUTA' not in col_u) or col_u == 'HAS'):
            renombres[col] = 'AREA_MAESTRA'
            area_ok = True
            
    df.rename(columns=renombres, inplace=True)
    return df
    
def convertir_pesos(val):
    try:
        v = str(val)
        v_limpio = "".join([c for c in v if c.isdigit() or c in ['.', ',']])
        v_limpio = v_limpio.rstrip('.,')
        if v_limpio == '': return 0.0
        
        if ',' in v_limpio and '.' not in v_limpio: v_limpio = v_limpio.replace(',', '.')
        elif '.' in v_limpio and ',' in v_limpio: v_limpio = v_limpio.replace('.', '').replace(',', '.')
        elif '.' in v_limpio:
            partes = v_limpio.split('.')
            if len(partes[-1]) == 3: v_limpio = v_limpio.replace('.', '')
                
        num = float(v_limpio)
        if 0 < num < 2000: num = num * 1000 
        return num
    except: return 0.0

def limpiar_area(val):
    try:
        v = str(val).upper().replace(',', '.')
        v = "".join([c for c in v if c.isdigit() or c == '.'])
        return float(v) if v != '' else 0.0
    except: return 0.0

def calcular_frecuencia(df):
    if df.empty or 'FECHA_DT' not in df.columns: return 0, 0
    fechas = sorted(df['FECHA_DT'].dt.date.unique())
    if not fechas: return 0, 0
    
    ciclos = 1
    inicios_ciclo = [fechas[0]]
    for i in range(1, len(fechas)):
        if (fechas[i] - fechas[i-1]).days > 5:
            ciclos += 1
            inicios_ciclo.append(fechas[i])
            
    if ciclos > 1:
        diffs =
