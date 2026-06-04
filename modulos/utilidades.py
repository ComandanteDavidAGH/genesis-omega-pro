import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime
import dateutil.parser

# =================================================================
# ⚡ PATRONES DE BÚSQUEDA PRE-COMPILADOS (ULTRA VELOCIDAD EN RAM)
# =================================================================
# Compilar los patrones aquí arriba evita que Python gaste ciclos de CPU 
# compilándolos miles de veces dentro de los bucles de los DataFrames.
REGEX_LOTE = re.compile(r'[^A-Z0-9]')
REGEX_NUMERO = re.compile(r'[^\d.,-]')
REGEX_ANIO = re.compile(r'\d{4}')
REGEX_DIA = re.compile(r'\b\d{1,2}\b')

# =================================================================
# ⚙️ CAJA DE HERRAMIENTAS Y AUDITORÍA DE DATOS CRUDOS
# =================================================================

def purificar_lote(lote):
    """ Purifica y estandariza los códigos de lote eliminando caracteres raros """
    if pd.isna(lote) or lote is None: 
        return ""
    return REGEX_LOTE.sub('', str(lote).upper().strip())

def quitar_tildes(s):
    """ Remueve de forma atómica acentos y tildes para evitar desajustes de strings """
    if pd.isna(s) or s is None: 
        return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(s).upper().strip()) if unicodedata.category(c) != 'Mn')

def extraer_numero(valor):
    """ Conversor universal seguro: Extrae el valor numérico flotante real de Excel/SAP """
    if pd.isna(valor) or valor == "": 
        return 0.0
    if isinstance(valor, (int, float)): 
        return float(valor)
        
    v = str(valor).strip().upper().replace("$", "").replace(" ", "")
    v = REGEX_NUMERO.sub('', v)
    
    # Cruce de formatos de miles y decimales
    if '.' in v and ',' in v: 
        v = v.replace('.', '').replace(',', '.')
    elif ',' in v: 
        v = v.replace(',', '.')
    elif v.count('.') > 1: 
        # 🎯 AJUSTE: Si hay múltiples puntos y ninguna coma (ej: 1.250.000), es formato de miles puro
        v = v.replace('.', '')
        
    try: 
        return float(v)
    except: 
        return 0.0

def fmt_sap(val): 
    """ Formatea números al estándar visual de SAP con puntos de miles """
    try:
        return f"{int(round(val, 0)):,}".replace(",", ".")
    except:
        return "0"

def limpiar_texto_vba(t):
    """ Sincroniza y limpia los textos heredados de macros VBA antiguas """
    if t is None: 
        return ""
    temp = str(t).upper().strip()
    temp = temp.replace(chr(160), " ").replace(".", "")
    
    # ⚡ CORRECCIÓN CRÍTICA: Reemplaza dobles espacios por espacios simples de forma segura.
    # La versión anterior generaba un bucle infinito que congelaba la CPU.
    while "  " in temp: 
        temp = temp.replace("  ", " ")
    return temp

def val_seguro(v):
    """ Encapsulador rápido try-except para flotantes en matrices lógicas """
    try: 
        return float(v)
    except: 
        return 0.0

def limpiar_val_dom(v):
    """ Decodifica de forma segura los valores de recargos dominicales """
    if v is None: 
        return 0.0
    s = str(v).strip()
    if s in ["", "-", "0", "0.0"]: 
        return 0.0 
    try:
        s = s.replace('$', '').replace(' ', '').replace(',', '.')
        return float(s)
    except: 
        return 0.0

def procesar_fecha_pesada(v):
    """ Radar Cronológico: Traduce formatos mixtos, números seriales de Excel y texto a datetime """
    if not v or str(v).strip() == "": 
        return None
        
    try:
        # Caso A: El valor es un número serial puro de Excel (int o float)
        if isinstance(v, (int, float)):
            f = datetime(1899, 12, 30) + pd.Timedelta(days=int(v))
            return f if f.year > 2020 else None
            
        v_str = str(v).lower().strip()
        
        # Caso B: El valor es un string numérico serial de Excel (ej: "45120")
        if v_str.replace('.', '').isdigit():
            f = datetime(1899, 12, 30) + pd.Timedelta(days=int(float(v_str)))
            return f if f.year > 2020 else None
            
        # Caso C: Formato verbal en español latino (ej: "lunes, enero 10, 2026")
        meses = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }
        for mes, num_mes in meses.items():
            if mes in v_str:
                match_ano = REGEX_ANIO.search(v_str)
                match_dia = REGEX_DIA.search(v_str)
                if match_ano and match_dia:
                    f = datetime(int(match_ano.group()), num_mes, int(match_dia.group()))
                    return f if f.year > 2020 else None
                    
        # Caso D: Formatos estándar de barra o guion (ej: 10/05/2026)
        if "/" in v_str or "-" in v_str:
            f = dateutil.parser.parse(v_str, dayfirst=True)
            return f if f.year > 2020 else None
            
    except: 
        pass
    return None
