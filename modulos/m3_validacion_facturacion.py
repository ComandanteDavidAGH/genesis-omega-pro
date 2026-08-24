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

# =================================================================
# 🛩️ MOTOR DEL SIMULADOR PRINCIPAL
# =================================================================
def ejecutar(procesar_fecha_pesada, extraer_numero):
    VERDE_INTENSO = '#143521'
    DORADO = '#d4af37'
    hora_oficial_col = obtener_hora_colombia()
    hoy_colombia_date = hora_oficial_col.date()

    st.markdown(f"""
    <style>
    .titulo-simulador {{ color: #0d1b2a; border-bottom: 3px solid {DORADO}; padding-bottom: 5px; font-family: 'Arial Black'; }}
    [data-testid="column"] {{
        display: flex !important;
        flex-direction: column !important;
        justify-content: flex-start !important;
        align-items: stretch !important;
    }}
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {{
        border: 3px solid #0d1b2a !important; 
        border-radius: 8px !important; 
        overflow: hidden !important; 
        box-shadow: 0px 4px 10px rgba(0,0,0,0.08) !important;
    }}
    div[data-testid="stSelectbox"] > div,
    div[data-testid="stDateInput"] > div,
    div[data-testid="stNumberInput"] > div,
    div[data-testid="stTextInput"] > div {{
        border: 2px solid {VERDE_INTENSO} !important;
        border-radius: 8px !important;
        background-color: #ffffff !important;
        box-shadow: 0px 4px 8px rgba(0,0,0,0.06) !important;
        overflow: hidden !important;
    }}
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    div[data-testid="stDateInput"] div[data-baseweb="input"],
    div[data-testid="stNumberInput"] div[data-baseweb="input"],
    div[data-testid="stTextInput"] div[data-baseweb="input"] {{
        background-color: transparent !important;
        border: none !important;
    }}
    div[data-testid="stSelectbox"] *,
    div[data-testid="stDateInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input {{
        color: #0d1b2a !important;
        font-weight: 900 !important;
    }}
    div[data-testid="stDateInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input {{
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }}
    div[data-testid="stMainBlockContainer"] label p {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    c_t, c_btn = st.columns([3, 1])
    with c_t:
        st.markdown("<h1 class='titulo-simulador'>Análisis de Validación y Facturación</h1>", unsafe_allow_html=True)
    with c_btn:
        st.write("")
        if st.button("🔄 Sincronizar Módulo", type="primary", use_container_width=True, key="btn_sync_m3"):
            st.cache_data.clear()
            st.session_state.fecha_sim_mem = hoy_colombia_date
            if 'fecha_vuelo_master' in st.session_state: del st.session_state['fecha_vuelo_master']
            st.toast("✅ Módulo Sincronizado y Memoria Vaciada.", icon="🔄")
            st.rerun()

    with st.container(border=True):
        st.markdown("### 📡 Panel de Operaciones")
        c_finca, c_pedido, c_fecha = st.columns([2, 2, 1.3])
        if 'fecha_sim_mem' not in st.session_state: st.session_state.fecha_sim_mem = hoy_colombia_date

        fecha_operacion = c_fecha.date_input("📅 Fecha de Vuelo", value=st.session_state.fecha_sim_mem, format="DD/MM/YYYY", key="fecha_vuelo_master")

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
        finca_sel = c_finca.selectbox("📍 Seleccione Finca:", opciones_finca)

        if 'finca_anterior' not in st.session_state: st.session_state.finca_anterior = finca_sel
        if 'fecha_operacion_anterior' not in st.session_state: st.session_state.fecha_operacion_anterior = fecha_operacion

        if (finca_sel != st.session_state.finca_anterior) or (fecha_operacion != st.session_state.fecha_operacion_anterior):
            st.session_state.dias_ciclo_sim_mem = calcular_dias_ciclo_real(finca_sel, fecha_operacion)
            st.session_state.finca_anterior = finca_sel
            st.session_state.fecha_operacion_anterior = fecha_operacion
            st.rerun()

        dias_ciclo_calc = calcular_dias_ciclo_real(finca_sel, fecha_operacion)

        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros Base e Inteligencia de Ciclos")
            
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            with r1c1:
                st.number_input("📅 CICLO (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key="ciclo_sistema_key")
            with r1c2:
                st.number_input("⏳ CICLO (COBRO)", value=int(dias_ciclo_calc), step=1, key="ciclo_cobro_key")
            with r1c3:
                st.number_input("🧪 HA DOSIS (TOTAL 459)", value=29.17, key="ha_dosis_key")
            with r1c4:
                st.markdown("<div style='margin-top:25px;'></div>", unsafe_allow_html=True)
                st.caption("🔒 Ciclos Verificados en Tiempo Real")

if __name__ == "__main__":
    pass
