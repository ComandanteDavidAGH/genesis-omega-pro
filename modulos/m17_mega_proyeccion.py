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

# 🛰️ ENLACES NATIVOS
from modulos.utilidades import procesar_fecha_pesada
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =================================================================
# 🔌 CONEXIÓN Y MOTORES DE FORMATO REGIONAL BLINDADOS
# =================================================================

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

def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

def limpiar_tarifa_excel(val):
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip().replace("$", "").replace(" ", "").upper()
    if not v or v in ['-', 'NAN', 'NONE', '']: return 0.0
    
    s_clean = re.sub(r'[^\d\.,\-]', '', v)
    try:
        if '.' in s_clean and ',' in s_clean:
            if s_clean.rfind(',') > s_clean.rfind('.'): s_clean = s_clean.replace('.', '').replace(',', '.')
            else: s_clean = s_clean.replace(',', '')
        elif ',' in s_clean:
            if len(s_clean.split(',')[-1]) == 3: s_clean = s_clean.replace(',', '')
            else: s_clean = s_clean.replace(',', '.')
        elif '.' in s_clean:
            if s_clean.count('.') > 1: s_clean = s_clean.replace('.', '')
            elif len(s_clean.split('.')[-1]) == 3: s_clean = s_clean.replace('.', '')
        return float(s_clean) if s_clean else 0.0
    except:
        return 0.0

def normalizar_a_fecha_pura(val):
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)): return res_nativo.date()
        if isinstance(res_nativo, date): return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except: return None

@st.cache_data(show_spinner=False, ttl=60)
def cargar_bases_m17(url_boveda, url_precios, supabase_client=None):
    gc = obtener_cliente_gspread_unificado()
    if not gc: return None, None, None, None, None, pd.DataFrame()
    
    df_mezclas, df_conf, df_dicc, df_t2, df_precios, df_t1 = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try:
        boveda_recetas = gc.open_by_url(url_boveda)
        sh_precios = gc.open_by_url(url_precios)
        
        # 1. BÓVEDA LIGERA
        try: df_mezclas = pd.DataFrame(boveda_recetas.worksheet("DD_Mesclas").get_all_values()[1:], columns=boveda_recetas.worksheet("DD_Mesclas").get_all_values()[0])
        except: pass
        if not df_mezclas.empty: df_mezclas['COCTEL_CLEAN'] = df_mezclas.iloc[:, 0].astype(str).str.upper().str.replace(" ", "")

        try: df_conf = pd.DataFrame(boveda_recetas.worksheet("Configuración").get_all_values()[1:], columns=boveda_recetas.worksheet("Configuración").get_all_values()[0])
        except: pass
        
        # 🛸 EXTRACCIÓN HÍBRIDA DE SIGLAS (SUPABASE FIRST)
        if supabase_client:
            try:
                res = supabase_client.table("DICCIONARIO_SIGLAS").select("*").execute()
                if res.data:
                    df_dicc = pd.DataFrame(res.data)
                    df_dicc.columns = [str(c).upper().strip() for c in df_dicc.columns]
            except:
                pass

        # Fallback de respaldo a Google Sheets si Supabase no devolvió datos
        if df_dicc.empty:
            try: 
                dicc_raw = boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()
                if dicc_raw:
                    df_dicc = pd.DataFrame(dicc_raw[1:], columns=[str(c).upper().strip() for c in dicc_raw[0]])
            except: 
                pass
        
        try: 
            t2_raw = boveda_recetas.worksheet("TABLA 2").get_all_values()
            idx_t2 = next((i for i, r in enumerate(t2_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 0)
            df_t2 = pd.DataFrame(t2_raw[idx_t2+1:], columns=[str(c).strip() for c in t2_raw[idx_t2]])
        except: pass

        # 2. FRANCOTIRADOR DE PRECIOS
        try:
            ws_datos = sh_precios.worksheet("DATOS") 
            datos_hoja = ws_datos.get_all_values()
            precios_consolidados = []
            
            if datos_hoja:
                idx_header, col_anio, col_prod = -1, -1, -1
                for i in range(min(10, len(datos_hoja))):
                    fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                    if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                        idx_header, col_anio, col_prod = i, fila_upper.index('AÑO'), fila_upper.index('PRODUCTO'); break
                
                if idx_header != -1:
                    for row in datos_hoja[idx_header+1:]:
                        if len(row) > max(col_anio, col_prod):
                            anio_str, str_prod = str(row[col_anio]).strip().upper(), str(row[col_prod]).strip().upper()
                            if anio_str and str_prod:
                                vals = []
                                for v in row[max(col_anio, col_prod) + 1:]:
                                    val_c = re.sub(r'[^\d\.,\-]', '', str(v).strip())
                                    if val_c:
                                        if '.' in val_c and ',' in val_c: val_c = val_c.replace('.', '').replace(',', '.') if val_c.rfind(',') > val_c.rfind('.') else val_c.replace(',', '')
                                        elif ',' in val_c: val_c = val_c.replace(',', '.')
                                        try: vals.append(float(val_c))
                                        except: pass
                                if vals: precios_consolidados.append({'AÑO': anio_str, 'PRODUCTO': str_prod, 'PRODUCTO_CLEAN': str_prod.replace(" ", ""), 'PRECIO_PROM': sum(vals)/len(vals)})
            df_precios = pd.DataFrame(precios_consolidados)
        except: pass

        # 3. EXTRAER TABLA 1 (CON RESPALDO DE COORDENADAS POR POSICIÓN)
        try:
            t1_raw = boveda_recetas.worksheet("TABLA 1").get_all_values()
            if t1_raw:
                idx_t1 = next((i for i, r in enumerate(t1_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 4)
                encabezados = [str(c).upper().strip() for c in t1_raw[idx_t1]]
                df_t1 = pd.DataFrame(t1_raw[idx_t1+1:], columns=encabezados)
                
                col_finca = next((c for c in encabezados if "FINCA" in c or "PROPIEDAD" in c), None)
                if not col_finca and len(encabezados) > 2: col_finca = encabezados[2]
                
                col_fecha = next((c for c in encabezados if "FECHA" in c or "DATE" in c), None)
                if not col_fecha and len(encabezados) > 7: col_fecha = encabezados[7]
                
                col_costo_ha = next((c for c in encabezados if "COSTO" in c and "AVI" in c and "$/HA" in c.replace(" ", "")), None)
                if not col_costo_ha: col_costo_ha = next((c for c in encabezados if "COSTO" in c and "$/HA" in c.replace(" ", "")), None)
                if not col_costo_ha and len(encabezados) > 19: col_costo_ha = encabezados[19]
                
                if col_finca and col_costo_ha:
                    def limp_num_col(val):
                        v = str(val).strip()
                        if not v or v == '-': return 0.0
                        v = re.sub(r'[^\d\.,\-]', '', v)
                        if not v: return 0.0
                        try:
                            if v.count('.') == 1 and v.count(',') == 0:
                                if len(v.split('.')[1]) == 3: v = v.replace('.', '')
                            if '.' in v and ',' in v:
                                if v.rfind(',') > v.rfind('.'): v = v.replace('.', '').replace(',', '.')
                                else: v = v.replace(',', '')
                            elif ',' in v: v = v.replace(',', '.')
                            f_val = float(v)
                            if f_val < 1000 and '.' in str(val) and len(str(val).split('.')[-1]) == 3:
                                f_val = f_val * 1000
                            return f_val
                        except: return 0.0
                    
                    df_t1['F_CLEAN'] = df_t1[col_finca].astype(str).apply(lambda x: re.sub(r'[^A-Z0-9]', '', x.upper().strip()))
                    df_t1['VAL_COSTO_HA'] = df_t1[col_costo_ha].apply(limp_num_col)
                    if col_fecha:
                        df_t1['FECHA_CLEAN'] = df_t1[col_fecha].astype(str).str.strip()
        except: pass
                            
    except Exception as e: 
        raise Exception(f"Error de conexión con Google Drive: {e}")

    return df_mezclas, df_conf, df_dicc, df_t2, df_precios, df_t1

# =================================================================
# 🧠 MOTORES DE LÓGICA Y EMPAREJAMIENTO INTELIGENTE
# =================================================================

def limpiar_numero(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip().replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1: v = v.rsplit('.', 1)[0].replace('.', '') + '.' + v.rsplit('.', 1)[1]
        return float(v) if v else 0.0
    except: return 0.0

def calcular_promedio_vuelo_finca(finca_usuario, df_t1):
    if df_t1 is None or df_t1.empty or 'VAL_COSTO_HA' not in df_t1.columns or 'F_CLEAN' not in df_t1.columns: 
        return 45000.0
    
    finca_buscada = re.sub(r'[^A-Z0-9]', '', str(finca_usuario).upper().strip())
    df_finca = df_t1[df_t1['F_CLEAN'] == finca_buscada]
    
    if df_finca.empty and finca_buscada:
        match_inicial = df_t1['F_CLEAN'].str.startswith(finca_buscada, na=False)
        df_finca = df_t1[match_inicial]
    
    if df_finca.empty: 
        return 45000.0 
        
    año_actual = str(datetime.now().year)
    año_corto = año_actual[-2:]
    
    if 'FECHA_CLEAN' in df_finca.columns:
        mask_año = df_finca['FECHA_CLEAN'].str.contains(año_actual, na=False) | df_finca['FECHA_CLEAN'].str.endswith(f"/{año_corto}", na=False) | df_finca['FECHA_CLEAN'].str.endswith(f"-{año_corto}", na=False)
        df_finca_año = df_finca[mask_año]
        
        if not df_finca_año.empty:
            df_valid_costos = df_finca_año[df_finca_año['VAL_COSTO_HA'] > 1000]
            if not df_valid_costos.empty:
                prom = df_valid_costos['VAL_COSTO_HA'].mean()
                return 45000.0 if pd.isna(prom) else float(prom)
    
    df_valid_costos_hist = df_finca[df_finca['VAL_COSTO_HA'] > 1000]
    if not df_valid_costos_hist.empty:
        prom = df_valid_costos_hist['VAL_COSTO_HA'].mean()
        return 45000.0 if pd.isna(prom) else float(prom)
            
    return 45000.0

def extraer_receta_mega(coctel_sel, finca_sel, df_mezclas, df_dicc, df_t2):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if partes else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    es_organico = False
    try:
        if not df_t2.empty:
            match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_sel.upper().strip()]
            if not match_f.empty and "ORGANIC" in str(match_f.iloc[0, 5]).upper(): es_organico = True
    except: pass

    base_buscar = f"{base_coctel}O" if es_organico and not base_coctel.endswith('O') else base_coctel

    if not df_mezclas.empty:
        col_0 = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0 == base_buscar]
        if rb.empty and es_organico: rb = df_mezclas[col_0 == base_coctel]
        for _, r in rb.iterrows():
            p, d = str(r.iloc[1]).strip().upper(), limpiar_numero(r.iloc[2])
            if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    if not df_dicc.empty and aditivos:
        for ad in aditivos:
            m_s = df_dicc[df_dicc['SIGLA'].astype(str).str.upper().str.strip() == ad]
            if not m_s.empty:
                p_ad, d_ad = str(m_s.iloc[0]['PRODUCTO']).strip().upper(), limpiar_numero(m_s.iloc[0]['DOSIS'])
                if d_ad > 0 and p_ad not in ['NAN', 'NONE', '']: dict_prods[p_ad] = dict_prods.get(p_ad, 0.0) + d_ad

    fert_fallback = {"ZN": "ZINTRAC X LITRO SV", "BT": "BANATREL SC", "NM": "NATURAMIN WSP", "QM": "QUELAMIX", "ZT": "ZITRON"}
    for ad in aditivos:
        if ad in fert_fallback:
            p_fall = fert_fallback[ad]
            if not any(p_fall in k for k in dict_prods.keys()):
                d_fall = 0.5 
                if not df_mezclas.empty:
                    try:
                        for col_idx in range(len(df_mezclas.columns) - 1):
                            mask = df_mezclas.iloc[:, col_idx].astype(str).str.strip().str.upper() == p_fall
                            if mask.any():
                                val = limpiar_numero(df_mezclas[mask].iloc[0, col_idx+1])
                                if val > 0: d_fall = val
                    except: pass
                dict_prods[p_fall] = dict_prods.get(p_fall, 0.0) + d_fall

    for p in list(dict_prods.keys()):
        if "ACONDICIONADOR" in p: dict_prods[p] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
        elif "IMBIOSIL" in p.replace(" ", ""): dict_prods[p] = 1.5 if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel else 1.0
        if es_organico and "ADHERENTE" in p: del dict_prods[p]
    if es_organico and not any("SPRAYFIX" in k for k in dict_prods.keys()): dict_prods["SPRAYFIX"] = 0.2
    
    return dict_prods

# =================================================================
# 👑 RENDERIZADO VISUAL PRINCIPAL
# =================================================================

def ejecutar(supabase_client=None):
    VERDE_INTENSO = '#143521'
    DORADO = '#d4af37'

    st.markdown(f"""
    <style>
    .titulo-mega {{ color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; margin-bottom: 15px;}}
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {{ border: 3px solid #143521 !important; border-radius: 8px !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); overflow: hidden !important; }}
    .tarjeta-kpi {{ background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.2); text-align: center; margin-bottom: 15px;}
    .kpi-titulo {{ font-size: 12px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }}
    .kpi-valor {{ font-size: 24px; font-family: 'Arial Black'; margin: 5px 0 0 0; }}
    
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] {{
        background-color: #ffffff !important;
        border: 3px solid {VERDE_INTENSO} !important;
        border-radius: 6px !important;
    }}
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {{
        background-color: transparent !important;
        border: none !important;
    }}
    div[data-testid="stTextInput"] *, div[data-testid="stNumberInput"] *, div[data-testid="stMultiSelect"] * {{
        color: #000000 !important;
        font-weight: bold !important;
    }}
    div[data-testid="stMainBlockContainer"] label p {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-mega'>🚀 Módulo 17: Mega-Proyección Operativa</h1>", unsafe_allow_html=True)

    db_cargada = st.session_state.get('m17_db_cargada', False)
    
    if db_cargada and 'm17_t1' not in st.session_state:
        st.session_state['m17_db_cargada'] = False
        db_cargada = False

    with st.expander("🔌 CONEXIÓN A LAS MAESTRAS DE GOOGLE DRIVE Y BASE DE DATOS", expanded=not db_cargada):
        if db_cargada:
            st.success("✅ Bases de Datos conectadas and en Memoria RAM del Módulo 17.")
        else:
            st.info("💡 Pega los enlaces de tus archivos de Google Sheets para alimentar la proyección.")
            
        url_1 = st.text_input("🔗 Link Bóveda (Recetas, Fincas, Tabla 1):", value=st.session_state.get('m17_url1', ''))
        url_2 = st.text_input("🔗 Link Comparativo de Precios:", value=st.session_state.get('m17_url2', ''))

        if st.button("🔄 Conectar y Descargar", type="primary"):
            if url_1 and url_2:
                with st.spinner("Descargando información (Modo Original Protegido)..."):
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
                        st.success("¡Extracción perfecta! Memoria cargada híbrida realizada.")
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))
            else:
                st.warning("⚠️ Debes pegar ambos enlaces.")

    if not db_cargada:
        st.stop() 

    df_mezclas = st.session_state.get('m17_mez', pd.DataFrame())
    df_conf = st.session_state.get('m17_conf', pd.DataFrame())
    df_dicc = st.session_state.get('m17_dicc', pd.DataFrame())
    df_t2 = st.session_state.get('m17_t2', pd.DataFrame())
    df_precios = st.session_state.get('m17_prec', pd.DataFrame())
    df_t1 = st.session_state.get('m17_t1', pd.DataFrame())

    columnas_base = ["FINCA", "HECTAREAS", "COCTEL", "FERTILIZANTE", "DIAS CICLO", "PRECIO VUELO"]
    
    if 'm17_df_entrada_grid' not in st.session_state:
        st.session_state.m17_df_entrada_grid = pd.DataFrame([{
            "FINCA": "", "HECTAREAS": "", "COCTEL": "", "FERTILIZANTE": "", "DIAS CICLO": "", "PRECIO VUELO": ""
        } for _ in range(500)])

    st.markdown("### 📥 1. Pista de Aterrizaje Segura")
    st.caption("📋 Selecciona tus columnas en Excel, haz Ctrl+C, párate en la primera celda de abajo y presiona **Ctrl+V**.")
    
    df_edited = st.data_editor(
        st.session_state.m17_df_entrada_grid,
        key="m17_tabla_maestra_grid", 
        use_container_width=True,
        hide_index=True,
        column_config={
            "FINCA": st.column_config.TextColumn("Finca"),
            "HECTAREAS": st.column_config.TextColumn("Hectáreas"), 
            "COCTEL": st.column_config.TextColumn("Cóctel"),
            "FERTILIZANTE": st.column_config.TextColumn("Fertilizante"),
            "DIAS CICLO": st.column_config.TextColumn("Días Ciclo"),
            "PRECIO VUELO": st.column_config.TextColumn("Precio Vuelo Manual (Opcional)"),
        }
    )

    st.markdown("---")
    st.markdown("### ⚙️ 2. Parámetros de Riesgo y Proyección")
    col_r1, col_r2 = st.columns(2)
    inflacion_proyectada = col_r1.number_input("📈 Inflación Global Proyectada (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
    colchon_dias = col_r2.number_input("🛡️ Colchón de Días Ciclo (Sumar a todas)", min_value=0, max_value=30, value=0, step=1)

    factor_inflacion = 1 + (inflacion_proyectada / 100)

    if st.button("🔥 EJECUTAR MEGA-PROYECCIÓN", type="primary", use_container_width=True):
        
        df_valid = df_edited.dropna(subset=['FINCA']).copy()
        df_valid = df_valid[df_valid['FINCA'].astype(str).str.strip() != ""]
        
        if df_valid.empty:
            st.error("⚠️ La tabla está vacía. Por favor pega datos antes de ejecutar.")
        else:
            with st.spinner("Procesando matriz financiera y logística..."):
                
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
                    ha_num = limpiar_numero(row['HECTAREAS'])
                    coctel_n = str(row['COCTEL']).strip().upper() if pd.notna(row['COCTEL']) else ""
                    
                    fert_n = str(row.get('FERTILIZANTE', '')).strip().upper() if pd.notna(row.get('FERTILIZANTE')) and str(row.get('FERTILIZANTE')).strip().upper() != "NONE" else ""
                    coctel_combinado = f"{coctel_n} {fert_n}".strip()

                    dias_c = int(limpiar_numero(row['DIAS CICLO'])) + colchon_dias
                    precio_vuelo_manual = limpiar_numero(row['PRECIO VUELO'])

                    if ha_num == 0 and not df_t2.empty:
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                        if not match_f.empty:
                            ha_num = limpiar_numero(match_f.iloc[0].iloc[2])

                    if ha_num <= 0: continue

                    if precio_vuelo_manual == 0:
                        precio_vuelo_final = calcular_promedio_vuelo_finca(finca_n, df_t1)
                    else:
                        precio_vuelo_final = precio_vuelo_manual

                    precio_vuelo_final = precio_vuelo_final * factor_inflacion

                    tipo_prod = "TERCERO"
                    if not df_t2.empty:
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                        if not match_f.empty: tipo_prod = str(match_f.iloc[0].iloc[col_prod_idx]).strip().upper() if len(match_f.columns) > col_prod_idx else "TERCERO"
                    
                    if "COOP" in finca_n or "EMPREBANCOOP" in finca_n: tipo_prod = "COOPERATIVA"

                    mult_m, st_base, mult_v = 1.112, 1337.0, 1.112
                    if not df_conf.empty:
                        match_cfg = df_conf[df_conf.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_prod]
                        if not match_cfg.empty:
                            mult_m = limpiar_numero(match_cfg.iloc[0].iloc[3])
                            st_base = limpiar_numero(match_cfg.iloc[0].iloc[4])
                            mult_v = limpiar_numero(match_cfg.iloc[0].iloc[6])
                    
                    if mult_m == 0 or st_base == 0:
                        if tipo_prod == "TERCERO": mult_m, st_base, mult_v = 1.451, 1583.0, 1.451
                        elif tipo_prod == "AFILIADO": mult_m, st_base, mult_v = 1.164, 1510.0, 1.164
                        elif tipo_prod == "COOPERATIVA": mult_m, st_base, mult_v = 1.112, 1510.0, 1.164
                        elif tipo_prod == "ORGANICO": mult_m, st_base, mult_v = 1.011, 1337.0, 1.011
                        else: mult_m, st_base, mult_v = 1.112, 1337.0, 1.112 

                    st_base = st_base * factor_inflacion

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
                        if not df_conf.empty:
                            mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.strip() == p
                            if not mask_cfg.any() and "NEMATICIDA" in p: mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.contains("NEMATI", na=False)
                            if mask_cfg.any(): precio_unitario = limpiar_numero(df_conf[mask_cfg].iloc[0, c_c_i])
                        
                        if precio_unitario == 0.0 and not df_precios.empty:
                            match_p = df_precios[df_precios['PRODUCTO_CLEAN'] == p.replace(" ","")]
                            if not match_p.empty: precio_unitario = match_p['PRECIO_PROM'].mean()

                        precio_unitario = precio_unitario * factor_inflacion
                        costo_mezcla_fila += (d * ha_num * precio_unitario * mult_m)

                    costo_st_fila = dias_c * st_base * ha_num
                    costo_vuelo_fila = precio_vuelo_final * ha_num 

                    costo_mezcla_fila = 0.0 if pd.isna(costo_mezcla_fila) else float(costo_mezcla_fila)
                    costo_st_fila = 0.0 if pd.isna(costo_st_fila) else float(costo_st_fila)
                    costo_vuelo_fila = 0.0 if pd.isna(costo_vuelo_fila) else float(costo_vuelo_fila)

                    gran_total = math.floor(costo_mezcla_fila + costo_st_fila + costo_vuelo_fila + 0.5)
                    costo_ha = math.floor((gran_total / ha_num) + 0.5) if ha_num > 0 else 0

                    resultados.append({
                        "FINCA": finca_n, "HECTAREAS": ha_num, "COCTEL": coctel_combinado, "DIAS CICLO": dias_c, "PRECIO VUELO": precio_vuelo_final,
                        "Costo ST ($)": math.floor(costo_st_fila), "Costo Vuelo ($)": math.floor(costo_vuelo_fila), "Costo Mezcla ($)": math.floor(costo_mezcla_fila),
                        "Costo x Ha ($)": costo_ha, "RESULTADO TOTAL ($)": gran_total
                    })

                st.session_state.m17_resultados = pd.DataFrame(resultados)
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
        t_mx = df_filtro['Costo Mezcla ($)'].sum()
        t_gr = df_filtro['RESULTADO TOTAL ($)'].sum()

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>👨‍🔬 Total Serv. Tec</p><p class='kpi-valor'>$ {formato_latino(t_st, 0)}</p></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>✈️ Total Vuelo</p><p class='kpi-valor'>$ {formato_latino(t_vu, 0)}</p></div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>🧪 Total Mezcla</p><p class='kpi-valor'>$ {formato_latino(t_mx, 0)}</p></div>", unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='tarjeta-kpi' style='border-left: 5px solid #00ff00;'><p class='kpi-titulo' style='color:#00ff00;'>🔥 GRAN TOTAL</p><p class='kpi-valor'>$ {formato_latino(t_gr, 0)}</p></div>", unsafe_allow_html=True)

        df_resumen_finca = df_filtro.groupby('FINCA', as_index=False)[
            ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']
        ].sum()

        tab1, tab2, tab3 = st.tabs(["📊 Detalles Económicos Fila x Fila", "📑 Resumen Ejecutivo por Finca", "📦 Auditoría Volumétrica de Insumos"])
        
        with tab1:
            df_view = df_filtro.copy()
            for col in ["PRECIO VUELO", "Costo ST ($)", "Costo Vuelo ($)", "Costo Mezcla ($)", "Costo x Ha ($)", "RESULTADO TOTAL ($)"]:
                df_view[col] = df_view[col].apply(lambda x: f"$ {formato_latino(x, 0)}")
            st.dataframe(df_view, use_container_width=True, hide_index=True)

        with tab2:
            if not df_resumen_finca.empty:
                df_resumen_view = df_resumen_finca.copy()
                for col in ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']:
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
                        title=f"Top 15 Insumos Proyectados"
                    )
                    fig.update_traces(textposition='outside', textfont_size=12)
                    fig.update_layout(
                        yaxis={'categoryorder':'total ascending'}, 
                        plot_bgcolor='rgba(0,0,0,0)', 
                        margin=dict(r=100),
                        hovermode="closest" 
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
            
            borde = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                ws.sheet_view.showGridLines = False
                
                max_r = ws.max_row
                max_c = ws.max_column
                
                column_headers = {}
                for col_idx in range(1, max_c + 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 20
                    header_val = ws.cell(row=1, column=col_idx).value
                    column_headers[col_idx] = str(header_val).upper() if header_val else ""

                for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
                    for cell in row:
                        cell.border = borde
                        if cell.row == 1:
                            cell.fill = fill_header
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(vertical='center')
                            
                            col_name = column_headers.get(cell.column, "")
                            
                            if isinstance(cell.value, (int, float)):
                                if "COSTO" in col_name or "PRECIO" in col_name or "RESULTADO" in col_name or "TOTAL" in col_name:
                                    cell.number_format = '"$" #,##0' 
                                elif "HECTAREAS" in col_name or "VOLUMEN" in col_name:
                                    cell.number_format = '#,##0.0'

        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL)",
            data=buffer.getvalue(),
            file_name=f"MegaProyeccion_Operativa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    pass
