import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import gspread
import io
import re
import math
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =================================================================
# ⚙️ CONSTANTES CENTRALIZADAS (ÚNICA FUENTE DE VERDAD)
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

# =================================================================
# ⚡ MOTOR DE CONEXIÓN UNIFICADO (V41)
# =================================================================
@st.cache_resource(show_spinner=False)
def obtener_cliente_gspread_unificado():
    """ Centraliza la autenticación unificada con Google Cloud una sola vez en RAM """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    if "gcp_credentials" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_credentials"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    try:
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

# =================================================================
# 🛡️ UTILIDADES DE PURIFICACIÓN Y FORMATO
# =================================================================
def limpiar_orden_extrema(val):
    if pd.isna(val) or str(val).strip() == "": return "SIN_ORDEN"
    v = str(val).upper().strip()
    v = re.sub(r'\s+', '', v) 
    if v.endswith('.0'): v = v[:-2] 
    return v

def limpiar_cantidad(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).replace(" ", "").strip()
        if "," in texto and "." in texto:
            if texto.rfind(".") > texto.rfind(","): texto = texto.replace(",", "")
            else: texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except Exception:
        return 0.0

def limpiar_moneda(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).upper().replace("$", "").replace("COP", "").replace(" ", "").strip()
        if "." in texto and "," in texto:
            if texto.rfind(".") > texto.rfind(","): texto = texto.replace(",", "")
            else: texto = texto.replace(".", "").replace(",", ".")
        else:
            sep = "." if "." in texto else ("," if "," in texto else None)
            if sep:
                if texto.count(sep) > 1:
                    texto = texto.replace(sep, "")
                elif len(texto.split(sep)[-1]) == 3: 
                    texto = texto.replace(sep, "")
                else: 
                    texto = texto.replace(sep, ".")
        return float(texto) if texto else 0.0
    except Exception:
        return 0.0

def parsear_fecha_robusta(val):
    if pd.isna(val) or str(val).strip() == "": return pd.NaT
    s = str(val).strip().lower()
    if s.isdigit(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(int(s), 'D')
    meses = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12}
    match1 = re.search(r'(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})', s)
    if match1:
        dia_str, mes_str, anio_str = match1.groups()
        if mes_str in meses: return pd.to_datetime(f"{anio_str}-{meses[mes_str]:02d}-{int(dia_str):02d}")
    match2 = re.search(r'([a-z]+)\s+(\d{1,2}),\s+(\d{4})', s)
    if match2:
        mes_str, dia_str, anio_str = match2.groups()
        if mes_str in meses: return pd.to_datetime(f"{anio_str}-{meses[mes_str]:02d}-{int(dia_str):02d}")
    try: 
        return pd.to_datetime(s.split(" ")[0], dayfirst=True, errors='coerce')
    except Exception: 
        return pd.NaT

def purificar_datos_vuelo(eq_raw, pista_raw):
    eq = str(eq_raw).upper()
    p = str(pista_raw).upper()
    if "DRON" in eq or "DRONE" in eq:
        if "DATAROT" in eq or "PLUC" in p: return "DRONE DATAROT", "PLUC"
        if "NORTE" in eq or "PDIV" in p: return "DRONE NORTE", "PDIV"
        if "AVIL" in eq or "TEHO" in p: return "DRONE AVIL", "TEHO"
        if "GENESYS" in eq or "LUCI" in p: return "DRONE GENESYS", "LUCI"
        return "DRONE GENESYS", "LUCI" 
    if "TRUSH" in eq or "THRUS" in eq or "OMANDER" in eq: return "THRUS SR2", "AEROPENORT"
    if "PAWNEE" in eq or "BRAVO" in eq or "PIPER PA 36" in eq: return "PIPER PA 36-375", "AEROPENORT"
    if "AIR TRACTOR" in eq or "TRACTOR" in eq or "TOR" in eq: return "AIR TRACTOR", "FUMIGARAY"
    if "CESSNA" in eq or "PIPER PA 25" in eq:
        if "ASA" in p or "ASA" in eq: return "CESSNA ASA", "ASA"
        if "FUMIGARAY" in p or "FUMIGARAY" in eq: return "CESSNA FUMIGARAY", "FUMIGARAY"
        return "CESSNA O PIPER PA 25", "AEROPENORT"
    return "IGNORAR", "IGNORAR"

def formato_latino(numero, decimales=0):
    if pd.isna(numero) or numero is None: return "0"
    try:
        num = float(numero)
        if num == 0: return "0"
        if decimales == 0: texto_us = f"{num:,.0f}"
        else: texto_us = f"{num:,.{decimales}f}"
        return texto_us.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0"

def obtener_hora_colombia():
    return datetime.utcnow() + timedelta(hours=-5)

# =================================================================
# 📦 EXTRACCIÓN DE DATOS BLINDADA E INTEGRACIÓN DE CONFIGURACIÓN
# =================================================================
@st.cache_data(show_spinner=False, ttl=10)
def obtener_historial_completo_ciclos_cached():
    df_t1, df_apoyo = pd.DataFrame(), pd.DataFrame()
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame(), pd.DataFrame()
    try:
        boveda = gc.open_by_url(URL_BOVEDA_MAESTRA)
        t1 = boveda.worksheet("TABLA 1").get_all_values()
        idx_t1 = 4
        for i in range(min(6, len(t1))):
            if "FINCA" in [str(x).upper() for x in t1[i]]:
                idx_t1 = i; break
        df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1]) if len(t1) > idx_t1 else pd.DataFrame()
        
        apoyo = boveda.worksheet("TABLA DE APOYO2023").get_all_values()
        idx_ap = 0
        for i in range(min(20, len(apoyo))):
            if any('FINCA' in str(c).upper() for c in apoyo[i]): 
                idx_ap = i; break
        df_apoyo = pd.DataFrame(apoyo[idx_ap+1:], columns=apoyo[idx_ap]) if len(apoyo) > idx_ap else pd.DataFrame()
        
        return df_t1, df_apoyo
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

def calcular_dias_ciclo_real(finca_nombre, fecha_vuelo):
    if not finca_nombre or finca_nombre == "---": return 14
    try:
        f_obj_alpha = re.sub(r'[^A-Z0-9]', '', str(finca_nombre).upper())
        df_viva, df_hist = obtener_historial_completo_ciclos_cached()
        fechas_encontradas = []

        def extraer_fechas_motor(df_temp):
            if df_temp.empty: return
            col_f = next((c for c in df_temp.columns if 'FINCA' in str(c).upper() or 'PROPIEDAD' in str(c).upper()), None)
            col_d = next((c for c in df_temp.columns if 'FECHA' in str(c).upper() or 'DATE' in str(c).upper()), None)
            if col_f and col_d:
                fincas_alpha = df_temp[col_f].astype(str).str.upper().apply(lambda x: re.sub(r'[^A-Z0-9]', '', x))
                mask = fincas_alpha == f_obj_alpha
                if not mask.any(): mask = fincas_alpha.apply(lambda x: f_obj_alpha in x if f_obj_alpha else False)
                if not mask.any():
                    partes = f_obj_alpha.replace("COOP", "").replace("BANAFRU", "").replace("ASO", "").replace("COOBAMAG", "").strip()
                    clave = partes[:8] if len(partes) > 8 else partes
                    mask = fincas_alpha.str.contains(clave, regex=False, na=False)
                df_fil = df_temp[mask]
                for d_raw in df_fil[col_d]:
                    fecha_valida = parsear_fecha_robusta(d_raw)
                    if pd.notna(fecha_valida): fechas_encontradas.append(fecha_valida.date())

        extraer_fechas_motor(df_viva)
        extraer_fechas_motor(df_hist)
        
        if fechas_encontradas:
            fecha_vuelo_date = fecha_vuelo if isinstance(fecha_vuelo, date) else pd.to_datetime(fecha_vuelo).date()
            fechas_validas = [f for f in fechas_encontradas if f < fecha_vuelo_date]
            if fechas_validas:
                fecha_max = max(fechas_validas)
                dias = (fecha_vuelo_date - fecha_max).days
                if 0 <= dias <= 365: return int(dias)
    except Exception:
        pass
    return 14

@st.cache_data(show_spinner=False, ttl=600)
def extraer_datos_boveda():
    gc = obtener_cliente_gspread_unificado()
    df_t1, df_t2 = pd.DataFrame(), pd.DataFrame()
    dict_tarifas_conf = {}
    if not gc: return df_t1, df_t2, dict_tarifas_conf
    
    try:
        boveda = gc.open_by_url(URL_BOVEDA_MAESTRA)
        
        try:
            t1 = boveda.worksheet("TABLA 1").get_all_values()
            idx_t1 = 4
            for i in range(min(8, len(t1))):
                fila_limpia = [str(x).upper().strip() for x in t1[i]]
                if "Nº ORDEN" in fila_limpia or "FINCA" in fila_limpia or "VALOR A FACTURAR" in "".join(fila_limpia):
                    idx_t1 = i
                    break
            df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1]) if len(t1) > idx_t1 else pd.DataFrame()
        except Exception: pass
        
        try:
            hojas = [ws.title for ws in boveda.worksheets()]
            nombre_t2 = "TABLA 2" if "TABLA 2" in hojas else hojas[1]
            t2 = boveda.worksheet(nombre_t2).get_all_values()
            df_t2 = pd.DataFrame(t2[1:], columns=t2[0]) if len(t2)>1 else pd.DataFrame()
        except Exception: pass

        try:
            if "Configuración" in [ws.title for ws in boveda.worksheets()]:
                conf_data = boveda.worksheet("Configuración").get_all_values()
                if len(conf_data) > 1:
                    df_conf = pd.DataFrame(conf_data[1:], columns=conf_data[0])
                    for _, row in df_conf.iterrows():
                        key_eq = str(row.iloc[0]).strip().upper()
                        val_m = limpiar_moneda(row.iloc[1]) if len(row) > 1 else 0.0
                        if key_eq and val_m > 0:
                            dict_tarifas_conf[key_eq] = val_m
        except Exception: pass
        
    except Exception: pass
    
    return df_t1, df_t2, dict_tarifas_conf

@st.cache_data(show_spinner=False, ttl=600)
def cargar_matriz_tarifas_mod3():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    try:
        sh = gc.open_by_url(URL_BOVEDA_MAESTRA)
        ws = sh.worksheet("MATRIZ_TARIFAS")
        datos = ws.get_all_values()
        if len(datos) > 1:
            df = pd.DataFrame(datos[1:], columns=datos[0])
            df = df.loc[:, df.columns.astype(str).str.strip() != '']
            df = df[df['PISTA'].str.strip() != '']
            return df
    except Exception: pass
    return pd.DataFrame()

def extraer_tarifas_dinamicas(df_tarifas, anio_str):
    DICT_AVIONES_DEFAULT = {
        "THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA 25": 3036525,
        "AIR TRACTOR": 4665109, "CESSNA ASA": 3768500, "CESSNA FUMIGARAY": 3065952
    }
    DICT_DRONES_DEFAULT = {
        "DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280
    }
    dict_av, dict_dr = {}, {}
    dict_topes = {"TOPE MAX GENERAL": {}, "TOPE SUR": {}, "TOPE PARCELA INTER < 20HA": {}}
    if df_tarifas.empty: return DICT_AVIONES_DEFAULT, DICT_DRONES_DEFAULT, dict_topes, None
    anios_disp = [str(c) for c in df_tarifas.columns if str(c).isdigit()]
    col_anio = anio_str if anio_str in anios_disp else (max(anios_disp) if anios_disp else None)
    if col_anio:
        for _, r in df_tarifas.iterrows():
            pista, equipo, tarifa_val = str(r.get('PISTA', '')).strip().upper(), str(r.get('EQUIPO_O_TOPE', '')).strip().upper(), limpiar_moneda(r[col_anio])
            if "TOPE MAX" in equipo: dict_topes["TOPE MAX GENERAL"][pista] = tarifa_val
            elif "TOPE SUR" in equipo: dict_topes["TOPE SUR"][pista] = tarifa_val
            elif "TOPE PARCELA" in equipo or "20HA" in equipo: dict_topes["TOPE PARCELA INTER < 20HA"][pista] = tarifa_val
            elif "DRON" in equipo or "DR5" in equipo or "DATAROT" in equipo or "GENESYS" in equipo:
                dict_dr[equipo if "DRON" in equipo else f"DRONE {equipo}"] = tarifa_val
            elif equipo not in ["", "NAN"]: dict_av[equipo] = tarifa_val
    if not dict_av: dict_av = DICT_AVIONES_DEFAULT
    if not dict_dr: dict_dr = DICT_DRONES_DEFAULT
    return dict_av, dict_dr, dict_topes, col_anio

@st.cache_data(show_spinner=False, ttl=1800)
def obtener_matriz_fija_cruda_v2():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return []
    try:
        boveda = gc.open_by_url(URL_BOVEDA_MAESTRA)
        return boveda.worksheet("DD_Mesclas").get_all_values()
    except Exception: return []

def obtener_dosis_global_robusta_v2(df_mez_dummy, nombre_producto_sap):
    nombre_clean = re.sub(r'[^A-Z0-9]', '', str(nombre_producto_sap).upper())
    if not nombre_clean: return 0.0
    datos_crudos = obtener_matriz_fija_cruda_v2()
    if not datos_crudos: return 0.0
    for fila in datos_crudos:
        for c_idx in range(len(fila) - 1):
            try:
                val_clean = re.sub(r'[^A-Z0-9]', '', str(fila[c_idx]).upper())
                if val_clean and len(val_clean) >= 4 and (nombre_clean in val_clean or val_clean in nombre_clean):
                    val_num = re.sub(r'[^\d.]', '', str(fila[c_idx + 1]).replace(",", "."))
                    if val_num and val_num != ".":
                        dosis = float(val_num)
                        if 0 < dosis < 100: return dosis
            except Exception: continue
    return 0.0

@st.cache_data(show_spinner=False, ttl=1800)
def cargar_diccionarios_crudos():
    datos = obtener_matriz_fija_cruda_v2()
    dict_recetas, dict_lideres, dict_fertilizantes = {}, {}, {}
    if not datos: return dict_recetas, dict_lideres, dict_fertilizantes
    f_col = -1
    for r in range(min(20, len(datos))):
        for c in range(len(datos[r])):
            if 'FERTILIZANTE' in str(datos[r][c]).upper(): f_col = c; break
        if f_col != -1: break
    if f_col != -1 and f_col + 1 < len(datos[0]):
        for r in range(1, len(datos)):
            if len(datos[r]) > f_col + 1:
                nf, sf = str(datos[r][f_col]).strip().upper(), str(datos[r][f_col+1]).strip().upper()
                if nf and nf not in ["", "NAN", "NONE", "FERTILIZANTES"] and sf:
                    dict_fertilizantes[nf.replace(" ", "")] = sf
    for r in range(1, len(datos)):
        fila = datos[r]
        if len(fila) >= 3:
            cid, p_tabla, d_str = str(fila[0]).strip().upper(), str(fila[1]).strip().upper(), str(fila[2]).replace(",", ".")
            if cid and p_tabla and cid != "FINCA":
                p_clean = p_tabla.replace(" ", "")
                num_str = re.sub(r'[^\d.]', '', d_str)
                d_tabla = float(num_str) if num_str and num_str != "." else 0.0
                es_lider = True if len(fila) >= 4 and str(fila[3]).strip().upper() == "X" else False
                if cid not in dict_recetas: dict_recetas[cid] = {}
                dict_recetas[cid][p_clean] = d_tabla
                if es_lider: dict_lideres[cid] = p_clean
    return dict_recetas, dict_lideres, dict_fertilizantes

def emparejar_coctel_ia(sap_dict_pista, coctel_piloto_base):
    dict_recetas, dict_lideres, dict_fertilizantes = cargar_diccionarios_crudos()
    coctel_base, dosis_oficiales_coctel, max_p = "SIN COINCIDENCIA", {}, -9999
    tiene_acond_06 = any("ZINTRAC" in k.upper() or "ZITRON" in k.upper() or "BANATREL" in k.upper() for k in sap_dict_pista.keys())

    for iter_id, receta in dict_recetas.items():
        puntaje = 0
        lider_db = dict_lideres.get(iter_id, "")
        if lider_db and not any(lider_db == k or (len(k)>=4 and lider_db in k) or (len(lider_db)>=4 and k in lider_db) for k in sap_dict_pista.keys()):
            puntaje -= 1000
        for p_receta, d_esperada in receta.items():
            match_receta, dose_matched, match_perfecto = False, False, False
            d_receta_esperada = d_esperada
            if "ACONDICIONADOR" in p_receta: d_receta_esperada = 0.06 if tiene_acond_06 else 0.02
            elif "ACEITE" in p_receta:
                for char in iter_id:
                    if char.isdigit(): d_receta_esperada = float(char); break
            elif "IMBIOSIL" in p_receta: d_receta_esperada = 1.5 if str(iter_id).startswith("IN") else 1.0
                
            for k_sap, d_sap in sap_dict_pista.items():
                if p_receta == k_sap or (len(k_sap) >= 4 and p_receta in k_sap) or (len(p_receta) >= 4 and k_sap in p_receta):
                    match_receta = True
                    error, tolerancia = abs(d_sap - d_receta_esperada), max(0.05, d_receta_esperada * 0.15)
                    if error <= 0.05: match_perfecto = True; dose_matched = True
                    elif error <= tolerancia: dose_matched = True
                    break
            if match_receta:
                puntaje += 100
                if match_perfecto: puntaje += 100
                elif dose_matched: puntaje += 40
                else: puntaje -= 100
            else: puntaje -= 100

        for k_sap in sap_dict_pista.keys():
            if not any(p == k_sap or (len(k_sap)>=4 and p in k_sap) or (len(p)>=4 and k_sap in p) for p in receta.keys()):
                if not any(f == k_sap or (len(k_sap)>=4 and f in k_sap) or (len(f)>=4 and k_sap in f) for f in dict_fertilizantes.keys()):
                    puntaje -= 100

        if coctel_piloto_base and iter_id == coctel_piloto_base: puntaje += 50
        if puntaje > max_p:
            max_p = puntaje
            coctel_base = iter_id
            dosis_oficiales_coctel = receta.copy()

    sigla_fertilizante = ""
    for k_sap in sap_dict_pista.keys():
        for f_name, f_sigla in dict_fertilizantes.items():
            if f_name == k_sap or (len(k_sap) >= 4 and f_name in k_sap) or (len(f_name) >= 4 and k_sap in f_name):
                if not ("IMBIOSIL" in f_name and str(coctel_base).startswith("IN")):
                    sigla_fertilizante = f" {f_sigla}"; break
        if sigla_fertilizante: break

    return coctel_base + sigla_fertilizante if coctel_base != "SIN COINCIDENCIA" else "SIN COINCIDENCIA", dosis_oficiales_coctel

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL MÓDULO 3 (FIRMA CORREGIDA Y UNIFICADA)
# =================================================================
def ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada, *args):
    hora_oficial_col = obtener_hora_colombia()
    hoy_colombia_date = hora_oficial_col.date()

    PISTAS_VALIDAS = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
    PISTAS_DISPONIBLES_MATRIZ = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"]

    def sync_pistas(src, tgt):
        st.session_state[tgt] = st.session_state[src]

    st.header("", anchor="inicio_modulo")
    
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    [data-testid="column"] { display: flex !important; flex-direction: column !important; justify-content: flex-start !important; align-items: stretch !important; }
    [data-testid="column"] > div { width: 100% !important; }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important; overflow: hidden !important; }
    div[data-testid="stSelectbox"] > div, div[data-testid="stTextInput"] > div, div[data-testid="stNumberInput"] > div, div[data-testid="stDateInput"] > div { background-color: #ffffff !important; border: 2px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 4px 8px rgba(0,0,0,0.06) !important; }
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stDateInput"] input { border: none !important; box-shadow: none !important; font-weight: 800 !important; color: #0d1b2a !important; background-color: transparent !important; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div { background-color: transparent !important; border: none !important; }
    div[data-testid="stSelectbox"] * { color: #0d1b2a !important; font-weight: bold !important; }
    div[data-testid="stMainBlockContainer"] label p { color: #0d1b2a !important; font-weight: 800 !important; text-transform: uppercase !important; }
    div[data-testid="stCodeBlock"], div[data-testid="stCodeBlock"] pre { background-color: #ffffff !important; border: 2px solid #0d1b2a !important; border-radius: 8px !important; }
    div[data-testid="stCodeBlock"] code { color: #0d1b2a !important; font-weight: 900 !important; font-size: 17px !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Análisis de Validación y Facturación</h1>", unsafe_allow_html=True)
    
    df_tarifas_maestras = cargar_matriz_tarifas_mod3()

    col_vacia, col_sync = st.columns([3, 1])
    if col_sync.button("🔄 Sincronizar Módulo", type="primary", use_container_width=True, key="btn_sync_m3"):
        st.cache_data.clear()
        st.session_state.fecha_sim_mem = hoy_colombia_date
        if 'fecha_vuelo_master' in st.session_state: del st.session_state['fecha_vuelo_master']
        st.toast("✅ Módulo 3 Sincronizado y Memoria Vaciada.", icon="🔄")
        st.rerun()

    with st.container(border=True):
        st.markdown("### 📡 Panel de Operaciones")
        c_vacio, c_radar = st.columns([2, 2])
        pedido_sap = c_radar.text_input("📦 Buscar por N° Pedido SAP (Opcional):", key="buscar_sap_mod3", placeholder="Ej: 170036035")

        finca_sap = ""
        st.session_state['ha_radar_sap'] = 0.0 

        if pedido_sap and 'df_pedidos' in st.session_state:
            df_p = st.session_state['df_pedidos']
            match_sap = df_p[df_p.astype(str).apply(lambda x: x.str.contains(str(pedido_sap).strip())).any(axis=1)]
            if not match_sap.empty:
                try:
                    col_finca = [c for c in df_p.columns if any(x in str(c).upper() for x in ['FINCA', 'CLIENTE', 'DESTINATARIO', 'NOMBRE', 'SOLICITANTE'])][0]
                    col_ha = [c for c in df_p.columns if 'CANT' in str(c).upper() or 'HECT' in str(c).upper()][0]
                    col_mat = [c for c in df_p.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper() or 'COD' in str(c).upper()][0]
                    
                    finca_sap = str(match_sap.iloc[0][col_finca]).strip().upper()
                    ha_correcta = 0.0
                    for _, fila_ped in match_sap.iterrows():
                        valor_material = str(fila_ped[col_mat]).strip()
                        if valor_material == "459" or valor_material.split(".")[0] == "459": 
                            ha_correcta = limpiar_cantidad(fila_ped[col_ha]); break
                    st.session_state['ha_radar_sap'] = ha_correcta if ha_correcta > 0 else limpiar_cantidad(match_sap.iloc[0][col_ha])
                    st.success(f"✅ **SAP CONFIRMADO:** {finca_sap} | {st.session_state['ha_radar_sap']} Ha")
                except Exception: pass

        # 🎯 FILA DE 3 COLUMNAS: FINCA, REFERENCIA PEDIDO/INFORME, FECHA DE VUELO
        c_finca, c_pedido, c_fecha = st.columns([2, 2, 1.3])
        if 'fecha_sim_mem' not in st.session_state: st.session_state.fecha_sim_mem = hoy_colombia_date

        fecha_operacion = c_fecha.date_input("📅 Fecha de Vuelo", value=st.session_state.fecha_sim_mem, format="DD/MM/YYYY", key="fecha_vuelo_master")

        anio_vuelo = str(fecha_operacion.year)
        dict_aviones, dict_drones, dict_topes_pista, col_anio_detectado = extraer_tarifas_dinamicas(df_tarifas_maestras, anio_vuelo)

        df_t2 = st.session_state.get('df_config', pd.DataFrame())
        col_prod_idx_op, col_tope_idx_op = 5, 6
        if not df_t2.empty:
            for i, col_name in enumerate(df_t2.columns):
                c_up = str(col_name).upper()
                if 'PROD' in c_up or 'TIPO' in c_up: col_prod_idx_op = i
                if 'TOPE' in c_up: col_tope_idx_op = i
            lista_fincas_raw = df_t2.iloc[:, 0].dropna().astype(str).str.strip().str.upper().unique().tolist()
            lista_fincas = sorted([f for f in lista_fincas_raw if f not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']])
        else:
            df_base_tmp, df_t2_tmp, _ = extraer_datos_boveda()
            if not df_t2_tmp.empty:
                st.session_state['df_config'] = df_t2_tmp
                lista_fincas = sorted([str(f).strip().upper() for f in df_t2_tmp.iloc[:, 0].dropna().unique() if str(f).strip().upper() not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']])
            else:
                lista_fincas = ["RAQUELITA"]
                
        opciones_finca = ["---"] + lista_fincas
        idx_finca = 0
        if finca_sap:
            for i, f in enumerate(opciones_finca):
                if f.upper() in finca_sap or finca_sap in f.upper(): idx_finca = i; break

        finca_sel = c_finca.selectbox("📍 Seleccione Finca:", opciones_finca, index=idx_finca)
        vuegos_informe = st.session_state.get('df_pistas', pd.DataFrame())
        lista_origenes = vuegos_informe['ORIGEN'].unique().tolist() if not vuegos_informe.empty else []
        
        # 📄 CASILLA DE REFERENCIA PEDIDO/INFORME RESTAURADA
        vuelo_ref = c_pedido.selectbox("📄 Referencia Pedido/Informe:", ["---"] + lista_origenes)

        if 'vuelo_ref_anterior' not in st.session_state: st.session_state.vuelo_ref_anterior = vuelo_ref
        if vuelo_ref != st.session_state.vuelo_ref_anterior:
            st.session_state.vuelo_ref_anterior = vuelo_ref
            st.session_state.fecha_sim_mem = hoy_colombia_date
            if 'fecha_vuelo_master' in st.session_state: del st.session_state['fecha_vuelo_master']
            st.rerun()

        if finca_sel == "---" or vuelo_ref == "---":
            st.info("⚠️ Seleccione Finca y Pedido para rugir motores.")
            st.stop()

        casilla_key = f"{finca_sel}_{vuelo_ref}_{fecha_operacion}"
        llave_sistema = f"sys_limpio_v2_{casilla_key}"
        llave_cobro = f"cob_limpio_v2_{casilla_key}"

        if 'finca_anterior' not in st.session_state: st.session_state.finca_anterior = finca_sel
        if 'fecha_operacion_anterior' not in st.session_state: st.session_state.fecha_operacion_anterior = fecha_operacion

        if (finca_sel != st.session_state.finca_anterior) or (fecha_operacion != st.session_state.fecha_operacion_anterior):
            st.session_state.dias_ciclo_sim_mem = calcular_dias_ciclo_real(finca_sel, fecha_operacion)
            st.session_state.finca_anterior = finca_sel
            st.session_state.fecha_operacion_anterior = fecha_operacion
            st.rerun()

        dias_ciclo_calc = calcular_dias_ciclo_real(finca_sel, fecha_operacion)

        datos_vuelo = vuegos_informe[vuegos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo.get('DATOS_FILA', {})

        ha_cobro_detectada = limpiar_cantidad(datos_raw.get(8, 0))
        ha_dosis_detectada = st.session_state.get('ha_radar_sap', 0.0)
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros Base e Inteligencia de Ciclos")
            
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            with r1c1:
                st.number_input("📅 CICLO (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key=llave_sistema)
            with r1c2:
                st.number_input("⏳ CICLO (COBRO)", value=int(dias_ciclo_calc), step=1, key=llave_cobro)
            with r1c3:
                st.number_input("🧪 HA DOSIS (TOTAL 459)", value=float(ha_dosis_detectada), key=f"ha_dosis_{casilla_key}")
            with r1c4:
                st.markdown("<div style='margin-top:25px;'></div>", unsafe_allow_html=True)
                st.caption("🔒 Ciclos Verificados en Tiempo Real")

if __name__ == "__main__":
    pass
