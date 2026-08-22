import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import gspread
import re
import math
import io
import openpyxl
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# 🛰️ ENLACES NATIVOS
from modulos.utilidades import procesar_fecha_pesada

# =================================================================
# ⚙️ CONSTANTES CENTRALIZADAS Y VALORES POR DEFECTO
# =================================================================
COLOR_NAVY = '#0d1b2a'
COLOR_DORADO = '#d4af37'
COLOR_VERDE = '#143521'

TARIFA_VUELO_DEFAULT = 45000.0

# Multiplicadores de respaldo para cálculo de mezclas y servicios por tipo de producto
MULTIPLICADORES_FALLBACK = {
    "TERCERO": {"mult_m": 1.451, "st_base": 1583.0, "mult_v": 1.451},
    "AFILIADO": {"mult_m": 1.164, "st_base": 1510.0, "mult_v": 1.164},
    "COOPERATIVA": {"mult_m": 1.112, "st_base": 1510.0, "mult_v": 1.164},
    "ORGANICO": {"mult_m": 1.011, "st_base": 1337.0, "mult_v": 1.011},
    "DEFAULT": {"mult_m": 1.112, "st_base": 1337.0, "mult_v": 1.112}
}

FERTILIZANTES_FALLBACK = {
    "ZN": "ZINTRAC X LITRO SV",
    "BT": "BANATREL SC",
    "NM": "NATURAMIN WSP",
    "QM": "QUELAMIX",
    "ZT": "ZITRON"
}

# =================================================================
# 🛡️ MOTOR ÚNICO DE SANITIZACIÓN NUMÉRICA Y FORMATO
# =================================================================
def a_numero_limpio(val):
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    
    v = str(val).strip().replace("$", "").replace(" ", "").upper()
    if not v or v in ['-', 'NAN', 'NONE', '']: return 0.0
    
    s_clean = re.sub(r'[^\d\.,\-]', '', v)
    if not s_clean: return 0.0
    
    try:
        if '.' in s_clean and ',' in s_clean:
            if s_clean.rfind(',') > s_clean.rfind('.'):
                s_clean = s_clean.replace('.', '').replace(',', '.')
            else:
                s_clean = s_clean.replace(',', '')
        elif ',' in s_clean:
            if len(s_clean.split(',')[-1]) == 3:
                s_clean = s_clean.replace(',', '')
            else:
                s_clean = s_clean.replace(',', '.')
        elif '.' in s_clean:
            if s_clean.count('.') > 1:
                s_clean = s_clean.replace('.', '')
            elif len(s_clean.split('.')[-1]) == 3:
                s_clean = s_clean.replace('.', '')
                
        f_val = float(s_clean)
        if f_val < 1000 and '.' in str(val) and len(str(val).split('.')[-1]) == 3:
            f_val *= 1000.0
        return f_val
    except Exception:
        return 0.0

def formato_latino(numero, decimales=0):
    if pd.isna(numero) or numero is None: return "0"
    try:
        num = float(numero)
        if num == 0: return "0"
        texto_us = f"{num:,.{decimales}f}" if decimales > 0 else f"{num:,.0f}"
        return texto_us.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"

def normalizar_a_fecha_pura(val):
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)): return res_nativo.date()
        if isinstance(res_nativo, date): return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except Exception:
        return None

# =================================================================
# ⚡ MOTOR DE CONEXIÓN UNIFICADO
# =================================================================
@st.cache_resource(show_spinner=False)
def obtener_cliente_gspread_unificado():
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
# 💾 EXTRACCIÓN Y PROCESAMIENTO DE BASES DE DATOS
# =================================================================
@st.cache_data(show_spinner=False, ttl=120)
def cargar_bases_m17(url_boveda, url_precios, _supabase_client=None):
    gc = obtener_cliente_gspread_unificado()
    if not gc: return None, None, None, None, None, pd.DataFrame()
    
    df_mezclas, df_conf, df_dicc, df_t2, df_precios, df_t1 = [pd.DataFrame() for _ in range(6)]

    try:
        boveda_recetas = gc.open_by_url(url_boveda)
        sh_precios = gc.open_by_url(url_precios)
        
        # 1. MEZCLAS
        try:
            raw_m = boveda_recetas.worksheet("DD_Mesclas").get_all_values()
            if raw_m:
                df_mezclas = pd.DataFrame(raw_m[1:], columns=[str(c).strip() for c in raw_m[0]])
                df_mezclas['COCTEL_CLEAN'] = df_mezclas.iloc[:, 0].astype(str).str.upper().str.replace(" ", "")
        except Exception: pass

        # 2. CONFIGURACIÓN
        try:
            raw_c = boveda_recetas.worksheet("Configuración").get_all_values()
            if raw_c:
                df_conf = pd.DataFrame(raw_c[1:], columns=[str(c).strip() for c in raw_c[0]])
        except Exception: pass
        
        # 3. DICCIONARIO SIGLAS (SUPABASE FIRST -> FALLBACK DRIVE)
        if _supabase_client:
            try:
                res = _supabase_client.table("DICCIONARIO_SIGLAS").select("*").execute()
                if res.data:
                    df_dicc = pd.DataFrame(res.data)
                    df_dicc.columns = [str(c).upper().strip() for c in df_dicc.columns]
            except Exception: pass

        if df_dicc.empty:
            try: 
                dicc_raw = boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()
                if dicc_raw:
                    df_dicc = pd.DataFrame(dicc_raw[1:], columns=[str(c).upper().strip() for c in dicc_raw[0]])
            except Exception: pass
        
        # 4. TABLA 2
        try: 
            t2_raw = boveda_recetas.worksheet("TABLA 2").get_all_values()
            if t2_raw:
                idx_t2 = next((i for i, r in enumerate(t2_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 0)
                df_t2 = pd.DataFrame(t2_raw[idx_t2+1:], columns=[str(c).strip() for c in t2_raw[idx_t2]])
        except Exception: pass

        # 5. PRECIOS HISTÓRICOS
        try:
            ws_datos = sh_precios.worksheet("DATOS") 
            datos_hoja = ws_datos.get_all_values()
            precios_consolidados = []
            if datos_hoja:
                idx_header, col_anio, col_prod = -1, -1, -1
                for i in range(min(10, len(datos_hoja))):
                    fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                    if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                        idx_header, col_anio, col_prod = i, fila_upper.index('AÑO'), fila_upper.index('PRODUCTO')
                        break
                
                if idx_header != -1:
                    for row in datos_hoja[idx_header+1:]:
                        if len(row) > max(col_anio, col_prod):
                            anio_str, str_prod = str(row[col_anio]).strip().upper(), str(row[col_prod]).strip().upper()
                            if anio_str and str_prod:
                                vals = [a_numero_limpio(v) for v in row[max(col_anio, col_prod) + 1:] if a_numero_limpio(v) > 0]
                                if vals:
                                    precios_consolidados.append({
                                        'AÑO': anio_str, 
                                        'PRODUCTO': str_prod, 
                                        'PRODUCTO_CLEAN': str_prod.replace(" ", ""), 
                                        'PRECIO_PROM': sum(vals)/len(vals)
                                    })
            df_precios = pd.DataFrame(precios_consolidados)
        except Exception: pass

        # 6. TABLA 1 (HISTÓRICO OPERATIVO)
        try:
            t1_raw = boveda_recetas.worksheet("TABLA 1").get_all_values()
            if t1_raw:
                idx_t1 = next((i for i, r in enumerate(t1_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 4)
                encabezados = [str(c).upper().strip() for c in t1_raw[idx_t1]]
                df_t1 = pd.DataFrame(t1_raw[idx_t1+1:], columns=encabezados)
                
                col_finca = next((c for c in encabezados if "FINCA" in c or "PROPIEDAD" in c), None)
                if not col_finca and len(encabezados) > 2: col_finca = encabezados[2]
                
                col_fecha = next((c for c in encabezados if "FECHA" in c or "DATE" in c), None)
                col_costo_ha = next((c for c in encabezados if "COSTO" in c and "AVI" in c and "$/HA" in c.replace(" ", "")), None)
                if not col_costo_ha: col_costo_ha = next((c for c in encabezados if "COSTO" in c and "$/HA" in c.replace(" ", "")), None)
                col_recargo = next((c for c in encabezados if "DOMINIC" in c or "RECARGO" in c), None)
                
                if col_finca and col_costo_ha:
                    df_t1['F_CLEAN'] = df_t1[col_finca].astype(str).apply(lambda x: re.sub(r'[^A-Z0-9]', '', x.upper().strip()))
                    df_t1['VAL_COSTO_HA'] = df_t1[col_costo_ha].apply(a_numero_limpio)
                    df_t1['VAL_RECARGO_HA'] = df_t1[col_recargo].apply(a_numero_limpio) if col_recargo else 0.0
                    if col_fecha: df_t1['FECHA_CLEAN'] = df_t1[col_fecha].astype(str).str.strip()
        except Exception: pass
                            
    except Exception as e: 
        raise Exception(f"Error al procesar bases de datos: {e}")

    return df_mezclas, df_conf, df_dicc, df_t2, df_precios, df_t1

# =================================================================
# 🧠 CÁLCULOS HISTÓRICOS Y EXTRACCIÓN DE RECETAS
# =================================================================
def calcular_historicos_finca(finca_usuario, df_t1):
    if df_t1 is None or df_t1.empty or 'VAL_COSTO_HA' not in df_t1.columns or 'F_CLEAN' not in df_t1.columns: 
        return TARIFA_VUELO_DEFAULT, 0.0
    
    finca_buscada = re.sub(r'[^A-Z0-9]', '', str(finca_usuario).upper().strip())
    df_finca = df_t1[df_t1['F_CLEAN'] == finca_buscada]
    
    if df_finca.empty and finca_buscada:
        df_finca = df_t1[df_t1['F_CLEAN'].str.startswith(finca_buscada, na=False)]
    
    if df_finca.empty: 
        return TARIFA_VUELO_DEFAULT, 0.0 
        
    año_actual = str(datetime.now().year)
    año_corto = año_actual[-2:]
    
    df_evaluar = df_finca
    if 'FECHA_CLEAN' in df_finca.columns:
        mask_año = (df_finca['FECHA_CLEAN'].str.contains(año_actual, na=False) | 
                    df_finca['FECHA_CLEAN'].str.endswith(f"/{año_corto}", na=False) | 
                    df_finca['FECHA_CLEAN'].str.endswith(f"-{año_corto}", na=False))
        df_finca_año = df_finca[mask_año]
        if not df_finca_año.empty and not df_finca_año[df_finca_año['VAL_COSTO_HA'] > 1000].empty:
            df_evaluar = df_finca_año
            
    prom_vuelo = TARIFA_VUELO_DEFAULT
    prom_recargo = 0.0
            
    df_valid_costos = df_evaluar[df_evaluar['VAL_COSTO_HA'] > 1000]
    if not df_valid_costos.empty:
        prom_vuelo = float(df_valid_costos['VAL_COSTO_HA'].mean())
        if pd.isna(prom_vuelo): prom_vuelo = TARIFA_VUELO_DEFAULT

    if 'VAL_RECARGO_HA' in df_evaluar.columns:
        df_recargos_validos = df_evaluar[df_evaluar['VAL_RECARGO_HA'] > 100]
        if not df_recargos_validos.empty:
            prom_recargo = float(df_recargos_validos['VAL_RECARGO_HA'].mean())
            if pd.isna(prom_recargo): prom_recargo = 0.0
            
    return prom_vuelo, prom_recargo

def extraer_receta_mega(coctel_sel, finca_sel, df_mezclas, df_dicc, df_t2):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if partes else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    es_organico = False
    
    try:
        if not df_t2.empty and len(df_t2.columns) > 5:
            match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_sel.upper().strip()]
            if not match_f.empty and "ORGANIC" in str(match_f.iloc[0, 5]).upper(): 
                es_organico = True
    except Exception: pass

    base_buscar = f"{base_coctel}O" if es_organico and not base_coctel.endswith('O') else base_coctel

    if not df_mezclas.empty:
        col_0 = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0 == base_buscar]
        if rb.empty and es_organico: rb = df_mezclas[col_0 == base_coctel]
        for _, r in rb.iterrows():
            if len(r) >= 3:
                p, d = str(r.iloc[1]).strip().upper(), a_numero_limpio(r.iloc[2])
                if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    if not df_dicc.empty and aditivos:
        for ad in aditivos:
            if 'SIGLA' in df_dicc.columns and 'PRODUCTO' in df_dicc.columns and 'DOSIS' in df_dicc.columns:
                m_s = df_dicc[df_dicc['SIGLA'].astype(str).str.upper().str.strip() == ad]
                if not m_s.empty:
                    p_ad = str(m_s.iloc[0]['PRODUCTO']).strip().upper()
                    d_ad = a_numero_limpio(m_s.iloc[0]['DOSIS'])
                    if d_ad > 0 and p_ad not in ['NAN', 'NONE', '']: 
                        dict_prods[p_ad] = dict_prods.get(p_ad, 0.0) + d_ad

    for ad in aditivos:
        if ad in FERTILIZANTES_FALLBACK:
            p_fall = FERTILIZANTES_FALLBACK[ad]
            if not any(p_fall in k for k in dict_prods.keys()):
                d_fall = 0.5 
                if not df_mezclas.empty:
                    try:
                        for col_idx in range(len(df_mezclas.columns) - 1):
                            mask = df_mezclas.iloc[:, col_idx].astype(str).str.strip().str.upper() == p_fall
                            if mask.any():
                                val = a_numero_limpio(df_mezclas[mask].iloc[0, col_idx+1])
                                if val > 0: d_fall = val
                    except Exception: pass
                dict_prods[p_fall] = dict_prods.get(p_fall, 0.0) + d_fall

    for p in list(dict_prods.keys()):
        if "ACONDICIONADOR" in p: 
            dict_prods[p] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
        elif "IMBIOSIL" in p.replace(" ", ""): 
            dict_prods[p] = 1.5 if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel else 1.0
        if es_organico and "ADHERENTE" in p: 
            del dict_prods[p]
            
    if es_organico and not any("SPRAYFIX" in k for k in dict_prods.keys()): 
        dict_prods["SPRAYFIX"] = 0.2
    
    return dict_prods

# =================================================================
# 👑 RENDERIZADO PRINCIPAL
# =================================================================
def ejecutar(supabase_client=None):

    # CSS NATIVO VIP
    st.markdown(f"""
    <style>
    .titulo-mega {{ color: #0d1b2a; border-bottom: 3px solid {COLOR_DORADO}; padding-bottom: 5px; font-weight: 900; margin-bottom: 15px; text-transform: uppercase; }}
    
    /* CONFIGURACIÓN DEL EDITOR DE DATOS Y ENTRADAS */
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {{
        border: 2px solid {COLOR_NAVY} !important;
        border-radius: 8px !important;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.08) !important;
        overflow: hidden !important;
    }}
    
    div[data-testid="stTextInput"] > div, div[data-testid="stNumberInput"] > div {{
        background-color: #ffffff !important;
        border: 2px solid {COLOR_NAVY} !important;
        border-radius: 6px !important;
    }}
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
    }}
    
    .tarjeta-kpi {{
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 4px solid {COLOR_DORADO};
        padding: 12px;
        border-radius: 8px;
        color: white;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.15);
        text-align: center;
        margin-bottom: 15px;
    }}
    .kpi-titulo {{ font-size: 11px; font-weight: 800; color: {COLOR_DORADO}; text-transform: uppercase; margin:0; letter-spacing: 0.5px; }}
    .kpi-valor {{ font-size: 19px; font-weight: 900; margin: 4px 0 0 0; }}
    
    div[data-testid="stMainBlockContainer"] label p {{ color: #0d1b2a !important; font-weight: 800 !important; text-transform: uppercase !important; font-size: 12px !important; }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-mega'>🚀 Módulo 17: Mega-Proyección Operativa</h1>", unsafe_allow_html=True)

    db_cargada = st.session_state.get('m17_db_cargada', False)
    if db_cargada and 'm17_t1' not in st.session_state:
        st.session_state['m17_db_cargada'] = False
        db_cargada = False

    with st.expander("🔌 CONEXIÓN Y MAESTRAS DE BASE DE DATOS", expanded=not db_cargada):
        if db_cargada:
            st.success("✅ Bases de Datos conectadas y listas en Memoria RAM.")
        else:
            st.info("💡 Proporcione las URLs de Google Sheets para alimentar los modelos de receta y precios.")
            
        url_1 = st.text_input("🔗 Link Bóveda (Recetas, Fincas, Tabla 1):", value=st.session_state.get('m17_url1', ''))
        url_2 = st.text_input("🔗 Link Comparativo de Precios:", value=st.session_state.get('m17_url2', ''))

        if st.button("🔄 Conectar y Cargar Bases", type="primary"):
            if url_1 and url_2:
                with st.spinner("Sincronizando información de la Bóveda Maestra..."):
                    try:
                        mez, conf, dicc, t2, prec, t1 = cargar_bases_m17(url_1, url_2, supabase_client)
                        st.session_state['m17_mez'] = mez
                        st.session_state['m17_conf'] = conf
                        st.session_state['m17_dicc'] = dicc
                        st.session_state['m17_t2'] = t2
                        st.session_state['m17_prec'] = prec
                        st.session_state['m17_t1'] = t1
                        st.session_state['m17_url1'] = url_1
                        st.session_state['m17_url2'] = url_2
                        st.session_state['m17_db_cargada'] = True
                        st.success("¡Sincronización exitosa!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error de conexión: {e}")
            else:
                st.warning("⚠️ Debe ingresar ambos enlaces.")

    if not db_cargada:
        st.stop() 

    df_mezclas = st.session_state.get('m17_mez', pd.DataFrame())
    df_conf = st.session_state.get('m17_conf', pd.DataFrame())
    df_dicc = st.session_state.get('m17_dicc', pd.DataFrame())
    df_t2 = st.session_state.get('m17_t2', pd.DataFrame())
    df_precios = st.session_state.get('m17_prec', pd.DataFrame())
    df_t1 = st.session_state.get('m17_t1', pd.DataFrame())
    
    # 🎯 CORRECCIÓN DE RENDIMIENTO: Inicialización ligera con filas dinámicas (Evita freeze del DOM)
    if 'm17_df_entrada_grid' not in st.session_state or 'DOMINICAL' not in st.session_state.m17_df_entrada_grid.columns:
        st.session_state.m17_df_entrada_grid = pd.DataFrame([{
            "FINCA": "", "HECTAREAS": "", "COCTEL": "", "FERTILIZANTE": "", "DIAS CICLO": "", "PRECIO VUELO": "", "DOMINICAL": False
        } for _ in range(15)])

    st.markdown("### 📥 1. Entrada de Datos Operativos (Pegado masivo desde Excel)")
    st.caption("Copie sus columnas en Excel y péguelas directamente en la tabla. Use el botón '+' para agregar más filas si lo requiere.")
    
    df_edited = st.data_editor(
        st.session_state.m17_df_entrada_grid,
        key="m17_tabla_maestra_grid", 
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic", # Agregado dinámico para optimizar memoria
        column_config={
            "FINCA": st.column_config.TextColumn("Finca"),
            "HECTAREAS": st.column_config.TextColumn("Hectáreas"), 
            "COCTEL": st.column_config.TextColumn("Cóctel"),
            "FERTILIZANTE": st.column_config.TextColumn("Fertilizante"),
            "DIAS CICLO": st.column_config.TextColumn("Días Ciclo"),
            "PRECIO VUELO": st.column_config.TextColumn("Precio Vuelo Manual (Opcional)"),
            "DOMINICAL": st.column_config.CheckboxColumn("¿Dom/Fest?", default=False),
        }
    )

    st.markdown("---")
    st.markdown("### ⚙️ 2. Parámetros de Riesgo e Inflación")
    col_r1, col_r2 = st.columns(2)
    inflacion_proyectada = col_r1.number_input("📈 Inflación Global Proyectada (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
    colchon_dias = col_r2.number_input("🛡️ Colchón de Días Ciclo (Sumar a todas)", min_value=0, max_value=30, value=0, step=1)

    factor_inflacion = 1 + (inflacion_proyectada / 100)

    if st.button("🚀 EJECUTAR MEGA-PROYECCIÓN FINANCIAL", type="primary", use_container_width=True):
        
        df_valid = df_edited.dropna(subset=['FINCA']).copy()
        df_valid = df_valid[df_valid['FINCA'].astype(str).str.strip() != ""]
        
        if df_valid.empty:
            st.error("⚠️ La tabla de entrada no contiene registros válidos.")
        else:
            with st.spinner("Procesando matriz financiera y volumetría de insumos..."):
                
                col_prod_idx = 5
                if not df_t2.empty:
                    for i, c_name in enumerate(df_t2.columns):
                        c_clean = str(c_name).upper().replace('\n', ' ').strip()
                        if 'TIPO' in c_clean and 'PROD' in c_clean:
                            col_prod_idx = i
                            break 
                
                resultados = []
                log_volumetrico = {}

                for idx, row in df_valid.iterrows():
                    finca_n = str(row['FINCA']).strip().upper()
                    ha_num = a_numero_limpio(row['HECTAREAS'])
                    coctel_n = str(row['COCTEL']).strip().upper() if pd.notna(row['COCTEL']) else ""
                    
                    fert_n = str(row.get('FERTILIZANTE', '')).strip().upper() if pd.notna(row.get('FERTILIZANTE')) and str(row.get('FERTILIZANTE')).strip().upper() != "NONE" else ""
                    coctel_combinado = f"{coctel_n} {fert_n}".strip()

                    dias_c = int(a_numero_limpio(row['DIAS CICLO'])) + colchon_dias
                    precio_vuelo_manual = a_numero_limpio(row['PRECIO VUELO'])
                    aplica_dominical = bool(row.get('DOMINICAL', False))

                    if ha_num == 0 and not df_t2.empty and len(df_t2.columns) > 2:
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                        if not match_f.empty:
                            ha_num = a_numero_limpio(match_f.iloc[0].iloc[2])

                    if ha_num <= 0: continue

                    precio_vuelo_historico, recargo_historico = calcular_historicos_finca(finca_n, df_t1)

                    precio_vuelo_final = precio_vuelo_historico if precio_vuelo_manual == 0 else precio_vuelo_manual
                    precio_vuelo_final *= factor_inflacion
                    
                    recargo_final_ha = (recargo_historico * factor_inflacion) if aplica_dominical else 0.0

                    tipo_prod = "TERCERO"
                    if not df_t2.empty:
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                        if not match_f.empty and len(match_f.columns) > col_prod_idx:
                            tipo_prod = str(match_f.iloc[0].iloc[col_prod_idx]).strip().upper()
                    
                    if "COOP" in finca_n or "EMPREBANCOOP" in finca_n: tipo_prod = "COOPERATIVA"

                    # Búsqueda de multiplicadores o fallback
                    cfg_target = MULTIPLICADORES_FALLBACK.get(tipo_prod, MULTIPLICADORES_FALLBACK["DEFAULT"])
                    mult_m, st_base, mult_v = cfg_target["mult_m"], cfg_target["st_base"], cfg_target["mult_v"]

                    if not df_conf.empty:
                        match_cfg = df_conf[df_conf.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_prod]
                        if not match_cfg.empty and len(match_cfg.columns) >= 7:
                            mult_m = a_numero_limpio(match_cfg.iloc[0].iloc[3]) or mult_m
                            st_base = a_numero_limpio(match_cfg.iloc[0].iloc[4]) or st_base
                            mult_v = a_numero_limpio(match_cfg.iloc[0].iloc[6]) or mult_v

                    st_base *= factor_inflacion

                    costo_mezcla_fila = 0.0
                    c_p_i, c_c_i = 8, 9
                    if not df_conf.empty:
                        for i in range(min(5, len(df_conf))):
                            r_c = [str(x).upper() for x in df_conf.iloc[i]]
                            if 'PRODUCTO' in r_c and 'COSTO' in r_c:
                                c_p_i, c_c_i = r_c.index('PRODUCTO'), r_c.index('COSTO')
                                break

                    dict_receta = extraer_receta_mega(coctel_combinado, finca_n, df_mezclas, df_dicc, df_t2)
                    
                    for p, d in dict_receta.items():
                        log_volumetrico[finca_n] = log_volumetrico.get(finca_n, {})
                        log_volumetrico[finca_n][p] = log_volumetrico[finca_n].get(p, 0.0) + (d * ha_num)

                        precio_unitario = 0.0
                        if not df_conf.empty and len(df_conf.columns) > max(c_p_i, c_c_i):
                            mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.strip() == p
                            if not mask_cfg.any() and "NEMATICIDA" in p: 
                                mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.contains("NEMATI", na=False)
                            if mask_cfg.any(): 
                                precio_unitario = a_numero_limpio(df_conf[mask_cfg].iloc[0, c_c_i])
                        
                        if precio_unitario == 0.0 and not df_precios.empty:
                            match_p = df_precios[df_precios['PRODUCTO_CLEAN'] == p.replace(" ","")]
                            if not match_p.empty: 
                                precio_unitario = match_p['PRECIO_PROM'].mean()

                        precio_unitario *= factor_inflacion
                        costo_mezcla_fila += (d * ha_num * precio_unitario * mult_m)

                    costo_st_fila = float(dias_c * st_base * ha_num)
                    costo_vuelo_fila = float(precio_vuelo_final * ha_num)
                    costo_recargo_fila = float(recargo_final_ha * ha_num)
                    costo_mezcla_fila = float(costo_mezcla_fila)

                    gran_total = math.floor(costo_mezcla_fila + costo_st_fila + costo_vuelo_fila + costo_recargo_fila + 0.5)
                    costo_ha = math.floor((gran_total / ha_num) + 0.5) if ha_num > 0 else 0

                    resultados.append({
                        "FINCA": finca_n, 
                        "HECTAREAS": ha_num, 
                        "COCTEL": coctel_combinado, 
                        "DIAS CICLO": dias_c, 
                        "PRECIO VUELO": precio_vuelo_final, 
                        "RECARGO ($/HA)": recargo_final_ha,
                        "Costo ST ($)": math.floor(costo_st_fila), 
                        "Costo Vuelo ($)": math.floor(costo_vuelo_fila), 
                        "Costo Recargo ($)": math.floor(costo_recargo_fila), 
                        "Costo Mezcla ($)": math.floor(costo_mezcla_fila),
                        "Costo x Ha ($)": costo_ha, 
                        "RESULTADO TOTAL ($)": gran_total
                    })

                df_resultados_final = pd.DataFrame(resultados)
                if not df_resultados_final.empty:
                    df_resultados_final = df_resultados_final.sort_values(by="FINCA", ascending=True).reset_index(drop=True)

                st.session_state.m17_resultados = df_resultados_final
                st.session_state.m17_volumetria = log_volumetrico
                st.success("✅ Proyección completada exitosamente.")

    if 'm17_resultados' in st.session_state and not st.session_state.m17_resultados.empty:
        st.markdown("---")
        df_res = st.session_state.m17_resultados
        vol_dict = st.session_state.m17_volumetria
        
        fincas_procesadas = sorted(df_res['FINCA'].unique().tolist())
        
        st.markdown("### 🎛️ 3. Tablero de Mando y Filtros")
        fincas_filtro = st.multiselect("📍 Filtrar análisis por Finca(s) [Dejar vacío para ver TOTAL GENERAL]:", fincas_procesadas)
        
        if fincas_filtro:
            df_filtro = df_res[df_res['FINCA'].isin(fincas_filtro)]
            vol_dict_filtro = {k: v for k, v in vol_dict.items() if k in fincas_filtro}
        else:
            df_filtro = df_res
            vol_dict_filtro = vol_dict

        cons_vol_agrupado = {}
        for f, prods in vol_dict_filtro.items():
            for p, vol in prods.items():
                cons_vol_agrupado[p] = cons_vol_agrupado.get(p, 0.0) + vol
        
        t_st = df_filtro['Costo ST ($)'].sum()
        t_vu = df_filtro['Costo Vuelo ($)'].sum()
        t_re = df_filtro['Costo Recargo ($)'].sum() 
        t_mx = df_filtro['Costo Mezcla ($)'].sum()
        t_gr = df_filtro['RESULTADO TOTAL ($)'].sum()

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>👨‍🔬 Total Serv. Tec</p><p class='kpi-valor'>$ {formato_latino(t_st, 0)}</p></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>✈️ Total Vuelo</p><p class='kpi-valor'>$ {formato_latino(t_vu, 0)}</p></div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>⚠️ Total Recargos</p><p class='kpi-valor'>$ {formato_latino(t_re, 0)}</p></div>", unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>🧪 Total Mezcla</p><p class='kpi-valor'>$ {formato_latino(t_mx, 0)}</p></div>", unsafe_allow_html=True)
        with c5: st.markdown(f"<div class='tarjeta-kpi' style='border-left: 4px solid #28a745;'><p class='kpi-titulo' style='color:#28a745;'>🔥 GRAN TOTAL</p><p class='kpi-valor'>$ {formato_latino(t_gr, 0)}</p></div>", unsafe_allow_html=True)

        df_resumen_finca = df_filtro.groupby('FINCA', as_index=False)[
            ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Recargo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']
        ].sum()

        tab1, tab2, tab3 = st.tabs(["📊 Detalles Económicos Fila x Fila", "📑 Resumen Ejecutivo por Finca", "📦 Auditoría Volumétrica de Insumos"])
        
        with tab1:
            df_view = df_filtro.copy()
            for col in ["PRECIO VUELO", "RECARGO ($/HA)", "Costo ST ($)", "Costo Vuelo ($)", "Costo Recargo ($)", "Costo Mezcla ($)", "Costo x Ha ($)", "RESULTADO TOTAL ($)"]:
                df_view[col] = df_view[col].apply(lambda x: f"$ {formato_latino(x, 0)}")
            st.dataframe(df_view, use_container_width=True, hide_index=True)

        with tab2:
            if not df_resumen_finca.empty:
                df_resumen_view = df_resumen_finca.copy()
                for col in ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Recargo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']:
                    df_resumen_view[col] = df_resumen_view[col].apply(lambda x: f"$ {formato_latino(x, 0)}")
                st.dataframe(df_resumen_view, use_container_width=True, hide_index=True)
            else:
                st.info("No hay datos para resumir.")

        with tab3:
            if cons_vol_agrupado:
                df_insumos = pd.DataFrame(list(cons_vol_agrupado.items()), columns=["🧪 PRODUCTO", "VOLUMEN ESTIMADO"]).sort_values("VOLUMEN ESTIMADO", ascending=False)
                df_insumos["📦 VOLUMEN ESTIMADO (L/Kg)"] = df_insumos["VOLUMEN ESTIMADO"].apply(lambda x: formato_latino(x, 1))
                
                c_tbl, c_grf = st.columns([1, 1.2])
                with c_tbl:
                    st.dataframe(df_insumos[["🧪 PRODUCTO", "📦 VOLUMEN ESTIMADO (L/Kg)"]], use_container_width=True, hide_index=True)
                with c_grf:
                    df_grafica = df_insumos.head(15).copy()
                    fig = px.bar(
                        df_grafica, y="🧪 PRODUCTO", x="VOLUMEN ESTIMADO", text="📦 VOLUMEN ESTIMADO (L/Kg)",
                        orientation='h', color="VOLUMEN ESTIMADO", color_continuous_scale="GnBu",
                        title="Top 15 Insumos Proyectados"
                    )
                    fig.update_traces(textposition='outside', textfont_size=11)
                    fig.update_layout(
                        yaxis={'categoryorder':'total ascending'}, 
                        plot_bgcolor='rgba(0,0,0,0)', 
                        paper_bgcolor='rgba(0,0,0,0)',
                        margin=dict(r=80, t=30, l=10, b=10)
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No hay datos de insumos químicos para las fincas seleccionadas.")

        st.markdown("<br>", unsafe_allow_html=True)
        buffer = io.BytesIO()
        
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtro.to_excel(writer, sheet_name='Detalle_Económico', index=False)
            df_resumen_finca.to_excel(writer, sheet_name='Resumen_x_Finca', index=False)
            if cons_vol_agrupado:
                df_insumos[["🧪 PRODUCTO", "VOLUMEN ESTIMADO"]].to_excel(writer, sheet_name='Consumo_Insumos', index=False)
            
            workbook = writer.book
            borde = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), 
                           top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
            header_font = Font(color="D4AF37", bold=True)

            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                ws.sheet_view.showGridLines = True
                
                max_r = ws.max_row
                max_c = ws.max_column
                
                column_headers = {}
                for col_idx in range(1, max_c + 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 22
                    header_val = ws.cell(row=1, column=col_idx).value
                    column_headers[col_idx] = str(header_val).upper() if header_val else ""

                for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
                    for cell in row:
                        cell.border = borde
                        if cell.row == 1:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(vertical='center')
                            col_name = column_headers.get(cell.column, "")
                            if isinstance(cell.value, (int, float)):
                                if any(k in col_name for k in ["COSTO", "PRECIO", "RESULTADO", "TOTAL", "RECARGO"]):
                                    cell.number_format = '"$" #,##0' 
                                elif any(k in col_name for k in ["HECTAREAS", "VOLUMEN"]):
                                    cell.number_format = '#,##0.0'

        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL EN EXCEL",
            data=buffer.getvalue(),
            file_name=f"MegaProyeccion_Operativa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    pass
