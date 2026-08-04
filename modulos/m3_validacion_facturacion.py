import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import gspread
import requests
import io
import re
import math
import json
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 CONEXIÓN Y CARGA DE DATOS ULTRARRÁPIDA (SUPABASE FIRST)
# =================================================================

def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=1800)
def obtener_historial_completo_ciclos_cached():
    df_t1, df_apoyo = pd.DataFrame(), pd.DataFrame()
    
    if 'supabase' in st.session_state and st.session_state['supabase'] is not None:
        try:
            supabase_client = st.session_state['supabase']
            res_t1 = supabase_client.table("sap_tabla_1_maestro").select("*").execute()
            res_ap = supabase_client.table("tabla_apoyo_raw").select("*").execute()
            if res_t1.data: df_t1 = pd.DataFrame(res_t1.data)
            if res_ap.data: df_apoyo = pd.DataFrame(res_ap.data)
            if not df_t1.empty:
                return df_t1, df_apoyo
        except Exception:
            pass

    gc = obtener_cliente_gspread_unificado()
    if not gc:
        return pd.DataFrame(), pd.DataFrame()
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        
        t1 = boveda.worksheet("TABLA 1").get_all_values()
        idx_t1 = 4
        for i in range(min(6, len(t1))):
            if "FINCA" in [str(x).upper() for x in t1[i]]:
                idx_t1 = i
                break
        df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1]) if len(t1) > idx_t1 else pd.DataFrame()
        
        apoyo = boveda.worksheet("TABLA DE APOYO2023").get_all_values()
        idx_ap = 0
        for i in range(min(20, len(apoyo))):
            if any('FINCA' in str(c).upper() for c in apoyo[i]): 
                idx_ap = i
                break
        df_apoyo = pd.DataFrame(apoyo[idx_ap+1:], columns=apoyo[idx_ap]) if len(apoyo) > idx_ap else pd.DataFrame()
        
        return df_t1, df_apoyo
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data(show_spinner=False, ttl=1800)
def preprocesar_flota_gspread():
    gc = obtener_cliente_gspread_unificado()
    dict_aviones_default = {
        "THRUS SR2": 4606562, 
        "PIPER PA 36-375": 3985831, 
        "CESSNA O PIPER PA 25": 3036525, 
        "AIR TRACTOR": 4665109, 
        "CESSNA ASA": 3768500, 
        "CESSNA FUMIGARAY": 3065952
    }
    dict_drones_default = {
        "DRONE DATAROT": 84428, 
        "DRONE NORTE": 75518, 
        "DRONE AVIL": 71280, 
        "DRONE GENESYS": 71280
    }
    
    if not gc:
        return dict_aviones_default, dict_drones_default
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        datos_vd = boveda.worksheet("Validación Dosis").get_all_values()
        df_flota = pd.DataFrame(datos_vd[2:], columns=datos_vd[1])
        
        df_av = df_flota[df_flota['TIPO'].notna() & (df_flota['TIPO'].astype(str).str.strip() != '')]
        dict_aviones = dict(zip(df_av['TIPO'].astype(str).str.strip(), pd.to_numeric(df_av['HORA'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)))
        
        if "CESSNA ASA" in dict_aviones:
            dict_aviones["CESSNA ASA"] = 3768500
        
        df_dr = df_flota[df_flota['Tarifa'].notna() & (df_flota['Tarifa'].astype(str).str.strip() != '')]
        nombres_dr = df_dr['Tarifa'].astype(str).str.replace('TARIFA ', '', case=False).str.strip()
        nombres_dr = nombres_dr.apply(lambda x: f"DRONE {x}" if "DRONE" not in x.upper() else x)
        precios_dr = pd.to_numeric(df_dr['Valor ha/Dr'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)
        dict_drones = dict(zip(nombres_dr, precios_dr))
        
        return dict_aviones, dict_drones
    except Exception:
        return dict_aviones_default, dict_drones_default

# 💥 NUEVO CEREBRO CRUDO: Lee la base de datos evadiendo los errores de Pandas con las columnas vacías
@st.cache_data(show_spinner=False, ttl=1800)
def cargar_diccionarios_crudos():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return {}, {}, {}
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        datos = boveda.worksheet("DD_Mesclas").get_all_values()
        
        dict_recetas, dict_lideres, dict_fertilizantes = {}, {}, {}
        
        # Encontrar columna de fertilizantes
        f_col = -1
        for r in range(min(20, len(datos))):
            for c in range(len(datos[r])):
                if 'FERTILIZANTE' in str(datos[r][c]).upper():
                    f_col = c
                    break
            if f_col != -1: break
        
        # Poblar fertilizantes
        if f_col != -1 and f_col + 1 < len(datos[0]):
            for r in range(1, len(datos)):
                if len(datos[r]) > f_col + 1:
                    nf = str(datos[r][f_col]).strip().upper()
                    sf = str(datos[r][f_col+1]).strip().upper()
                    if nf and nf not in ["", "NAN", "NONE", "FERTILIZANTES"] and sf:
                        dict_fertilizantes[nf.replace(" ", "")] = sf

        # Construir recetas puras
        for r in range(1, len(datos)):
            fila = datos[r]
            if len(fila) >= 3:
                cid = str(fila[0]).strip().upper()
                p_tabla = str(fila[1]).strip().upper()
                d_str = str(fila[2]).replace(",", ".")
                
                if cid and p_tabla and cid != "FINCA":
                    p_clean = p_tabla.replace(" ", "")
                    num_str = re.sub(r'[^\d.]', '', d_str)
                    d_tabla = float(num_str) if num_str else 0.0
                    
                    es_lider = False
                    if len(fila) >= 4 and str(fila[3]).strip().upper() == "X":
                        es_lider = True
                        
                    if cid not in dict_recetas: dict_recetas[cid] = {}
                    dict_recetas[cid][p_clean] = d_tabla
                    if es_lider: dict_lideres[cid] = p_clean
                    
        return dict_recetas, dict_lideres, dict_fertilizantes
    except Exception:
        return {}, {}, {}

# 💥 SINERGIA: MOTOR CEREBRAL EXACTO DE TU MACRO "BUSCADOR_HIJO"
@st.cache_data(show_spinner=False, ttl=1800)
def emparejar_coctel_ia(sap_dict_pista, coctel_piloto_base):
    dict_recetas, dict_lideres, dict_fertilizantes = cargar_diccionarios_crudos()
    
    coctel_base = "SIN COINCIDENCIA"
    dosis_oficiales_coctel = {}
    max_p = -999

    for iter_id, receta in dict_recetas.items():
        es_valido = True
        puntaje = 0
        lider_db = dict_lideres.get(iter_id, "")
        
        # 1. VERIFICAR LÍDER DE LA RECETA
        match_lider = False
        if lider_db != "":
            for k_sap in sap_dict_pista.keys():
                k_sap_clean = k_sap.replace(" ", "")
                if lider_db == k_sap_clean or (len(k_sap_clean) >= 4 and lider_db in k_sap_clean) or (len(lider_db) >= 4 and k_sap_clean in lider_db) or (len(lider_db) >= 7 and len(k_sap_clean) >= 7 and k_sap_clean[:7] == lider_db[:7]):
                    match_lider = True
                    break
        
        if not match_lider and lider_db != "":
            es_valido = False # Guillotina: Faltó el líder
        else:
            puntaje += 1000

        # 2. VERIFICAR PRODUCTOS Y CALCULAR PUNTOS
        if es_valido:
            for p_receta, d_esperada in receta.items():
                d_receta_esperada = d_esperada
                
                # Reglas quemadas desde tu Excel
                if "ACONDICIONADOR" in p_receta:
                    if "ZN" in iter_id or "BT" in iter_id or "ZT" in iter_id or "ZITRON" in iter_id:
                        d_receta_esperada = 0.06
                    else:
                        d_receta_esperada = 0.02
                elif "ACEITE" in p_receta:
                    for char in iter_id:
                        if char.isdigit():
                            d_receta_esperada = float(char)
                            break
                elif "IMBIOSIL" in p_receta:
                    d_receta_esperada = 1.5 if str(iter_id).startswith("IN") else 1.0

                match_receta = False
                dose_matched = False
                
                for k_sap, d_sap in sap_dict_pista.items():
                    k_sap_clean = k_sap.replace(" ", "")
                    if p_receta == k_sap_clean or (len(k_sap_clean) >= 4 and p_receta in k_sap_clean) or (len(p_receta) >= 4 and k_sap_clean in p_receta) or (len(p_receta) >= 7 and len(k_sap_clean) >= 7 and k_sap_clean[:7] == p_receta[:7]):
                        match_receta = True
                        if abs(d_sap - d_receta_esperada) <= 0.2:
                            dose_matched = True
                        break
                
                if match_receta:
                    if dose_matched:
                        puntaje += 50
                    else:
                        puntaje += 10
                else:
                    es_valido = False # Guillotina: Faltó un ingrediente
                    break
        
        # 3. GUARDAR EL MEJOR PUNTAJE
        if es_valido:
            if puntaje > max_p:
                max_p = puntaje
                coctel_base = iter_id
                dosis_oficiales_coctel = receta.copy()

    # 4. RASTREO FERTILIZANTE FINAL (Como en la Macro)
    sigla_fertilizante = ""
    for k_sap in sap_dict_pista.keys():
        for f_name, f_sigla in dict_fertilizantes.items():
            if f_name == k_sap or (len(k_sap) >= 3 and f_name in k_sap) or (len(f_name) >= 3 and k_sap in f_name):
                if not ("IMBIOSIL" in f_name and str(coctel_base).startswith("IN")):
                    sigla_fertilizante = f" {f_sigla}"
                    break
        if sigla_fertilizante: 
            break

    final_coctel = coctel_base + sigla_fertilizante if coctel_base != "SIN COINCIDENCIA" else "SIN COINCIDENCIA"
    return final_coctel, dosis_oficiales_coctel

# =================================================================
# 👑 RENDERIZADO VISUAL PRINCIPAL
# =================================================================

def ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada):
    st.header("", anchor="inicio_modulo")

    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] { border: 3px solid #143521 !important; border-radius: 8px !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important; overflow: hidden !important; }
    div[data-testid="stSelectbox"] > div, div[data-testid="stSelectbox"] div[data-baseweb="select"], div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stDateInput"] input { background-color: #ffffff !important; border: 3px solid #143521 !important; border-radius: 8px !important; box-shadow: 0px 4px 8px rgba(0,0,0,0.06) !important; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div { background-color: transparent !important; border: none !important; }
    div[data-testid="stSelectbox"] *, div[data-testid="stTextInput"] *, div[data-testid="stNumberInput"] *, div[data-testid="stDateInput"] * { color: #000000 !important; font-weight: bold !important; }
    div[data-testid="stMainBlockContainer"] label p { color: #0d1b2a !important; font-weight: 800 !important; text-transform: uppercase !important; }
    div[data-testid="stCodeBlock"], div[data-testid="stCodeBlock"] pre, div[data-testid="stCodeBlock"] pre code { background-color: #ffffff !important; border: 3px solid #143521 !important; border-radius: 8px !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.08) !important; overflow: hidden !important; padding: 2px 5px !important; }
    div[data-testid="stCodeBlock"] code, div[data-testid="stCodeBlock"] code span, div[data-testid="stCodeBlock"] pre span { color: #0d1b2a !important; font-weight: 900 !important; font-size: 17px !important; font-family: 'Arial Black', monospace !important; }
    </style>
    """, unsafe_allow_html=True)

    def render_tarjetas_html(st_val, vuelo_val, mezcla_val, recargo_val, costo_ha_val):
        def f_h(val): return f"{val:,.0f}".replace(",", ".")
        return f"""
        <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; margin-bottom: 20px;">
            <div style="flex: 1; min-width: 120px; background-color: #ffffff; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; padding: 12px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.08);">
                <div style="font-size: 11px; color: #6c757d; font-weight: 800; text-transform: uppercase;">👨‍🔬 Serv. Tec</div>
                <div style="font-size: 17px; color: #0d1b2a; font-weight: 900; margin-top: 2px; user-select: all;" title="Doble clic para copiar">$ {f_h(st_val)}</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #ffffff; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; padding: 12px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.08);">
                <div style="font-size: 11px; color: #6c757d; font-weight: 800; text-transform: uppercase;">✈️ Vuelo</div>
                <div style="font-size: 17px; color: #0d1b2a; font-weight: 900; margin-top: 2px; user-select: all;" title="Doble clic para copiar">$ {f_h(vuelo_val)}</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #ffffff; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; padding: 12px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.08);">
                <div style="font-size: 11px; color: #6c757d; font-weight: 800; text-transform: uppercase;">🧪 Mezcla</div>
                <div style="font-size: 17px; color: #0d1b2a; font-weight: 900; margin-top: 2px; user-select: all;" title="Doble clic para copiar">$ {f_h(mezcla_val)}</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #ffffff; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; padding: 12px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.08);">
                <div style="font-size: 11px; color: #6c757d; font-weight: 800; text-transform: uppercase;">⚠️ Recargo</div>
                <div style="font-size: 17px; color: #0d1b2a; font-weight: 900; margin-top: 2px; user-select: all;" title="Doble clic para copiar">$ {f_h(recargo_val)}</div>
            </div>
            <div style="flex: 1.2; min-width: 140px; background-color: #0d1b2a; border: 3px solid #d4af37; padding: 12px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); text-align: center;">
                <div style="font-size: 11px; color: #d4af37; font-weight: 800; text-transform: uppercase;">💰 COSTO x HA</div>
                <div style="font-size: 19px; color: white; font-weight: 900; margin-top: 2px; user-select: all;" title="Doble clic para copiar">$ {f_h(costo_ha_val)}</div>
            </div>
        </div>
        """

    st.markdown("<h1 class='titulo-principal'>Análisis de Validación y Facturación</h1>", unsafe_allow_html=True)
    
    dict_aviones, dict_drones = preprocesar_flota_gspread()
    modo_simulacro = st.toggle("🔮 ACTIVAR MODO SIMULADOR (Modo Construcción de Matriz)")

    if modo_simulacro:
        st.info("💡 MODO CLON: Réplica exacta del Módulo de Validación.")
        st.stop()

    if 'df_pistas' not in st.session_state or 'df_apoyo' not in st.session_state:
        st.warning("🚨 No se detectan datos listos en el puente de mando.")
        st.info("💡 Por favor, cargue los tres archivos fuente en el Módulo 2 y procese antes de validar.")
        st.stop()

    def forzar_descarga_maestros():
        gc_maestro = obtener_cliente_gspread_unificado()
        if not gc_maestro: return None, None
        try:
            boveda_m = gc_maestro.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            t2_data = boveda_m.worksheet("TABLA 2").get_all_values()
            df_t2_temp = pd.DataFrame()
            if t2_data: 
                idx_t2 = 0
                for i in range(min(10, len(t2_data))):
                    if "FINCA" in [str(x).upper().strip() for x in t2_data[i]]:
                        idx_t2 = i
                        break
                cols_limpias = [str(c).strip() for c in t2_data[idx_t2]]
                df_t2_temp = pd.DataFrame(t2_data[idx_t2+1:], columns=cols_limpias)
                
            cfg_data = boveda_m.worksheet("Configuración").get_all_values()
            df_cfg_temp = pd.DataFrame(cfg_data[1:], columns=cfg_data[0]) if cfg_data else pd.DataFrame()
            return df_t2_temp, df_cfg_temp
        except Exception as e:
            st.error(f"🚨 Falla en la red: {e}")
            return None, None

    df_t2_cache = st.session_state.get('df_config', pd.DataFrame())
    df_cfg_cache = st.session_state.get('df_config_base', pd.DataFrame())

    necesita_sanacion = False
    if df_t2_cache.empty or df_cfg_cache.empty:
        necesita_sanacion = True
    else:
        fincas_reales = df_t2_cache.iloc[:, 0].dropna().astype(str).str.strip().str.upper().unique().tolist()
        fincas_reales = [f for f in fincas_reales if f not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']]
        if len(fincas_reales) < 5:
            necesita_sanacion = True  

    if necesita_sanacion:
        with st.spinner("🔄 Detectada anomalía de memoria. Iniciando Auto-Sanación de BD..."):
            df_t2_nueva, df_cfg_nueva = forzar_descarga_maestros()
            if df_t2_nueva is not None:
                st.session_state['df_config'] = df_t2_nueva
                st.session_state['df_config_base'] = df_cfg_nueva
                st.toast("✅ Base de Datos Sanada y Restaurada al 100%.", icon="🛠️")
            else:
                st.error("🚨 No se pudo restaurar la base de datos.") 

    col_vacia, col_sync = st.columns([3, 1])
    if col_sync.button("🔄 Sincronizar Módulo", type="primary", use_container_width=True, key="btn_sync_m3"):
        st.cache_data.clear()
        st.toast("✅ Módulo 3 Sincronizado y Memoria Vaciada.", icon="🔄")
        st.rerun()

    with st.container(border=True):
        st.markdown("### 📡 Panel de Operations")
        
        c_vacio, c_radar = st.columns([2, 2])
        pedido_sap = c_radar.text_input("📦 Buscar por N° Pedido SAP (Opcional):", key="buscar_sap_mod3", placeholder="Ej: 170036035")

        finca_sap = ""
        st.session_state['ha_radar_sap'] = 0.0 

        if pedido_sap and 'df_pedidos' in st.session_state:
            df_p = st.session_state['df_pedidos']
            match_sap = df_p[df_p.astype(str).apply(lambda x: x.str.contains(str(pedido_sap).strip())).any(axis=1)]
            
            if not match_sap.empty:
                try:
                    col_finca_cands = [c for c in df_p.columns if any(x in str(c).upper() for x in ['FINCA', 'CLIENTE', 'DESTINATARIO', 'NOMBRE', 'SOLICITANTE'])]
                    col_finca = col_finca_cands[0] if col_finca_cands else df_p.columns[8]
                    
                    col_ha_cands = [c for c in df_p.columns if 'CANT' in str(c).upper() or 'HECT' in str(c).upper()]
                    col_ha = col_ha_cands[0] if col_ha_cands else df_p.columns[6]
                    
                    col_mat_cands = [c for c in df_p.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper() or 'COD' in str(c).upper()]
                    col_mat = col_mat_cands[0] if col_mat_cands else df_p.columns[5]
                    
                    finca_sap = str(match_sap.iloc[0][col_finca]).strip().upper()
                    ha_correcta = 0.0
                    for _, fila_ped in match_sap.iterrows():
                        valor_material = str(fila_ped[col_mat]).strip()
                        if valor_material == "459" or valor_material.split(".")[0] == "459": 
                            ha_correcta = extraer_numero(fila_ped[col_ha])
                            break
                    
                    st.session_state['ha_radar_sap'] = ha_correcta if ha_correcta > 0 else extraer_numero(match_sap.iloc[0][col_ha])
                    st.success(f"✅ **SAP CONFIRMADO:** {finca_sap} | {st.session_state['ha_radar_sap']} Ha")
                except Exception: 
                    pass

        c0, c1, c2 = st.columns([1, 2, 2])
        fecha_operacion = c0.date_input("📅 Fecha de Vuelo", format="DD/MM/YYYY", key="fecha_vuelo_master")
        
        df_t2 = st.session_state.get('df_config', pd.DataFrame())
        
        col_prod_idx_op = 5
        col_tope_idx_op = 6
        if not df_t2.empty:
            for i, col_name in enumerate(df_t2.columns):
                c_up = str(col_name).upper()
                if 'PROD' in c_up or 'TIPO' in c_up: col_prod_idx_op = i
                if 'TOPE' in c_up: col_tope_idx_op = i
                
            lista_fincas_raw = df_t2.iloc[:, 0].dropna().astype(str).str.strip().str.upper().unique().tolist()
            lista_fincas = sorted([f for f in lista_fincas_raw if f not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']])
        else:
            st.error("🚨 CRÍTICO: El sistema no detecta la 'TABLA 2' maestra en memoria. Bloqueo de seguridad activado.")
            st.stop()
                
        opciones_finca = ["---"] + lista_fincas
        
        idx_finca = 0
        if finca_sap:
            for i, f in enumerate(opciones_finca):
                if f.upper() in finca_sap or finca_sap in f.upper(): 
                    idx_finca = i
                    break

        finca_sel = c1.selectbox("📍 Seleccione Finca:", opciones_finca, index=idx_finca)
        vuegos_informe = st.session_state.get('df_pistas', pd.DataFrame())
        lista_origenes = vuegos_informe['ORIGEN'].unique().tolist() if not vuegos_informe.empty else []
        vuelo_ref = c2.selectbox("📄 Referencia Pedido/Informe:", ["---"] + lista_origenes)

        if finca_sel == "---" or vuelo_ref == "---":
            st.info("⚠️ Seleccione Finca y Pedido para rugir motores.")
            st.stop()

        casilla_key = f"{finca_sel}_{vuelo_ref}_{fecha_operacion}"
        llave_sistema = f"sys_limpio_v2_{casilla_key}"
        llave_cobro = f"cob_limpio_v2_{casilla_key}"
        llave_editor_casilla = f"editor_valid_{casilla_key}"
        edited_df = pd.DataFrame()

        mult_material = 1.112
        tarifa_serv_tec_base = 1337.0
        mult_avion_base = 1.112
        
        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        
        finca_limpia = re.sub(r'\s+', ' ', str(finca_sel)).strip().upper()
        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        
        match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia]
        if not match_t2.empty:
            tipo_productor = str(match_t2.iloc[0].iloc[col_prod_idx_op]).strip().upper() if len(match_t2.columns) > col_prod_idx_op else "TERCERO"
            tipo_de_tope_finca = str(match_t2.iloc[0].iloc[col_tope_idx_op]).strip().upper() if len(match_t2.columns) > col_tope_idx_op else "SIN TOPE"

        if "COOP" in finca_limpia or "EMPREBANCOOP" in finca_limpia:
            tipo_productor = "COOPERATIVA"
        
        if not df_cfg.empty:
            match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
            if not match_cfg.empty:
                mult_material = extraer_numero(match_cfg.iloc[0].iloc[3])
                tarifa_serv_tec_base = extraer_numero(match_cfg.iloc[0].iloc[4])
                mult_avion_base = extraer_numero(match_cfg.iloc[0].iloc[6])

        dias_ciclo_calc = 0
        try:
            f_obj_alpha = re.sub(r'[^A-Z0-9]', '', finca_limpia)
            df_viva, df_hist = obtener_historial_completo_ciclos_cached()
            fechas_encontradas = []

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
                try: return pd.to_datetime(s.split(" ")[0], dayfirst=True, errors='coerce')
                except Exception: return pd.NaT

            def extraer_fechas(df_temp):
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
                        if pd.notna(fecha_valida): fechas_encontradas.append(fecha_valida)

            extraer_fechas(df_viva); extraer_fechas(df_hist)
            if fechas_encontradas:
                fecha_vuelo_dt = pd.to_datetime(fecha_operacion)
                fechas_validas = [f for f in fechas_encontradas if f < fecha_vuelo_dt]
                if fechas_validas:
                    fecha_max = max(fechas_validas)
                    dias_ciclo_calc = (fecha_vuelo_dt - fecha_max).days
                    if dias_ciclo_calc < 0 or dias_ciclo_calc > 365: dias_ciclo_calc = 0
        except Exception: pass

        datos_vuelo = vuegos_informe[vuegos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo.get('DATOS_FILA', {})
        
        num_pedido = "S/N"
        if pedido_sap and len(str(pedido_sap)) >= 7: num_pedido = str(pedido_sap).strip()
        elif datos_vuelo.get('PEDIDO_SAP') and str(datos_vuelo.get('PEDIDO_SAP')).strip() != "": num_pedido = str(datos_vuelo.get('PEDIDO_SAP')).strip()
        else:
            for idx in range(18, 40):
                val_celda = str(datos_raw.get(idx, "")).split('.')[0].strip()
                if val_celda.isdigit() and len(val_celda) >= 7: 
                    num_pedido = val_celda; break
        
        lista_pistas_validas = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
        pista_detectada = "PLUC"
        ha_dosis_detectada = 0.0
        match_ped = pd.DataFrame()

        if not df_ped.empty and num_pedido != "S/N":
            match_ped = df_ped[df_ped.astype(str).apply(lambda x: x.str.contains(num_pedido)).any(axis=1)]
            if not match_ped.empty:
                texto_pedido = match_ped.to_string().upper()
                for p_val in lista_pistas_validas:
                    if p_val in texto_pedido: pista_detectada = p_val; break
                col_ha_cands = [c for c in df_ped.columns if 'CANT' in str(c).upper() or 'HECT' in str(c).upper()]
                col_ha = col_ha_cands[0] if col_ha_cands else df_ped.columns[6]
                col_mat_cands = [c for c in df_ped.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper() or 'COD' in str(c).upper()]
                col_mat = col_mat_cands[0] if col_mat_cands else df_ped.columns[5]

                for _, r_p in match_ped.iterrows():
                    val_mat = str(r_p[col_mat]).strip()
                    if val_mat == "459" or val_mat.split(".")[0] == "459":
                        ha_dosis_detectada = extraer_numero(r_p[col_ha]); break
        
        ha_cobro_detectada = extraer_numero(datos_raw.get(8, 0))
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        coctel_piloto_raw = str(datos_vuelo.get('COCTEL', '')).upper().strip()
        partes_coctel = coctel_piloto_raw.replace("+", " ").replace("-", " ").split(" ")
        coctel_piloto_base = partes_coctel[0]

        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros Base e Inteligencia de Ciclos")
            c_sup1, c_sup2 = st.columns([3, 1])
            c_sup1.info(f"🧑‍🌾 Productor: **{tipo_productor}** | 🛣️ Tope: **{tipo_de_tope_finca}**")
            mision_solo_dron = c_sup2.toggle("🤖 MISIÓN 100% DRON", value=False, key=f"dron_toggle_{casilla_key}")
            
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            with r1c1: st.number_input("📅 Ciclo (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key=llave_sistema)
            with r1c2: d_ciclo_factura = st.number_input("⏳ Ciclo (COBRO)", value=int(dias_ciclo_calc), step=1, key=llave_cobro)
            with r1c3:
                ha_sugerida = float(st.session_state.get('ha_radar_sap', 0.0))
                if ha_sugerida == 0.0: ha_sugerida = float(ha_dosis_detectada)
                widget_key = f"had_{casilla_key}"
                sap_val = st.session_state.get('ha_radar_sap', 0.0)
                if sap_val > 0 and st.session_state.get(f"sync_{widget_key}") != sap_val:
                    st.session_state[widget_key] = float(sap_val)
                    st.session_state[f"sync_{widget_key}"] = sap_val
                ha_dosis_final = st.number_input("🧪 Ha Dosis (Total 459)", value=ha_sugerida, key=widget_key)
            with r1c4:
                multi_aviones = st.toggle("✈️ Recargo Coord. Multi-Avión", value=False, key=f"ma_{casilla_key}")
                multi_aviones_final = mult_avion_base + 0.1 if multi_aviones else mult_avion_base
                interciclo_menor_20 = st.toggle("🔄 Interciclo < 20ha", value=False, key=f"inter_{casilla_key}")

            recargo_final = 0.0
            pista_sel = "PLUC"
            if not mision_solo_dron:
                st.markdown("##### 🛣️ Parámetros Terrestres (Aviones)")
                r2c1, r2c2, r2c3 = st.columns(3)
                pista_sugerida = next((p for p in lista_pistas_validas if p in pista_detectada), "PLUC")
                pista_sel = r2c1.selectbox("Pista Base", lista_pistas_validas, index=lista_pistas_validas.index(pista_sugerida), key=f"pi_{casilla_key}")
                opciones_rec = ["0 (Sin Recargo)", "8740 (Porción PDIV)", "45000 (Recargo T. General)", "Otro Valor Manual..."]
                
                if f"pista_last_{casilla_key}" not in st.session_state:
                    st.session_state[f"pista_last_{casilla_key}"] = pista_sel
                    st.session_state[f"default_rec_idx_{casilla_key}"] = 1 if pista_sel == "PDIV" else 0
                elif st.session_state[f"pista_last_{casilla_key}"] != pista_sel:
                    st.session_state[f"pista_last_{casilla_key}"] = pista_sel
                    st.session_state[f"default_rec_idx_{casilla_key}"] = 1 if pista_sel == "PDIV" else 0
                    if f"rl_{casilla_key}" in st.session_state: del st.session_state[f"rl_{casilla_key}"]

                c_sap1, c_sap2, c_sap3, c_sap4 = st.columns(4) 
                recargo_lista = r2c2.selectbox("Cargo Terrestre:", opciones_rec, index=st.session_state[f"default_rec_idx_{casilla_key}"], key=f"rl_{casilla_key}")
                if recargo_lista == "Otro Valor Manual...": recargo_final = r2c3.number_input("✍️ Digite Recargo ($)", value=0, step=1000, key=f"rm_{casilla_key}")
                else: recargo_final = float(recargo_lista.split(" ")[0])

            dict_topes_pista = {
                "TOPE MAX GENERAL": {"PLUC": 63326, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325}, 
                "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517}, 
                "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}
            }
            
            tope_clave_efectiva = "TOPE PARCELA INTER < 20HA" if interciclo_menor_20 else tipo_de_tope_finca
            val_tope = dict_topes_pista.get(tope_clave_efectiva, {}).get(pista_sel, 999999)
            
            with st.container(border=True):
                st.markdown("#### ✈️ Hangar de Despliegue")
                costo_total_vuegos, costo_neto_vuelo_total, total_ha_cobro_escuadron, horometro_final_avion = 0.0, 0.0, 0.0, 0.0 

                if mision_solo_dron:
                    st.success("🚁 Modo Dron Activo: Costos calculados sin recargos terrestres ni topes de pista.")
                    df_drones_def = pd.DataFrame(columns=["Drone", "Hectáreas"])
                    escuadron_drones = st.data_editor(df_drones_def, key=f"drones_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)
                    for _, row in escuadron_drones.iterrows():
                        dr_sel, ha_dr = row.get("Drone"), row.get("Hectáreas")
                        if pd.isna(dr_sel) or ha_dr is None or float(ha_dr) <= 0: continue
                        ha_dr = float(ha_dr); total_ha_cobro_escuadron += ha_dr; tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                        costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                        costo_total_vuegos += (tarifa_dron_neta * ha_dr) * multi_aviones_final 
                else:
                    c_av, c_dr = st.columns(2)
                    with c_av: 
                        st.markdown("##### 🛩️ Base Aviones")
                        df_aviones_def = pd.DataFrame(columns=["Avión", "Hectáreas", "Horómetro"])
                        escuadron_aviones = st.data_editor(df_aviones_def, key=f"aviones_{casilla_key}", num_rows="dynamic", column_config={"Avión": st.column_config.SelectboxColumn("Modelo", options=list(dict_aviones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True), "Horómetro": st.column_config.NumberColumn("Horómetro", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)
                    with c_dr:
                        st.markdown("##### 🛸 Base Drones (Apoyo)")
                        df_drones_def = pd.DataFrame(columns=["Drone", "Hectáreas"])
                        escuadron_drones = st.data_editor(df_drones_def, key=f"drones_mix_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)                
                    for index, row in escuadron_aviones.iterrows():
                        av_sel, ha_av, horo = row.get("Avión"), row.get("Hectáreas"), row.get("Horómetro")
                        if pd.isna(av_sel) or ha_av is None or horo is None or float(ha_av) <= 0: continue
                        ha_av, horo = float(ha_av), float(horo); total_ha_cobro_escuadron += ha_av; horometro_final_avion += horo  
                        tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av if ha_av > 0 else 0
                        tarifa_base_tope = tarifa_base_ha if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope)
                        costo_neto_vuelo_total += (tarifa_base_tope * ha_av) 
                        costo_total_vuegos += ((tarifa_base_tope + recargo_final) * ha_av) * multi_aviones_final 
                    for _, row in escuadron_drones.iterrows():
                        dr_sel, ha_dr = row.get("Drone"), row.get("Hectáreas")
                        if pd.isna(dr_sel) or ha_dr is None or float(ha_dr) <= 0: continue
                        ha_dr = float(ha_dr); total_ha_cobro_escuadron += ha_dr; tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                        costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                        costo_total_vuegos += (tarifa_dron_neta * ha_dr) * multi_aviones_final
            
            st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Mezcla")
            pistas_disponibles = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"]
            pista_sel = st.selectbox("📍 Seleccione la Pista para extraer Inventario de SAP:", pistas_disponibles, index=pistas_disponibles.index(pista_sel), key="pista_matriz_maestra")
            
            st.markdown("---")
            costo_mezcla_total = 0.0

            if not match_ped.empty:
                idx_precio, idx_lote, idx_saldo, idx_almacen = -1, -1, -1, -1
                if not df_sab.empty:
                    for j, col in enumerate(df_sab.columns):
                        col_str = str(col).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').strip()
                        if ('MAYOR' in col_str or 'PRECIO' in col_str) and idx_precio == -1: idx_precio = j
                        if 'LOTE' in col_str and 'PROVEEDOR' not in col_str and idx_lote == -1: idx_lote = j
                        if ('ALMACEN' in col_str or 'PISTA' in col_str) and 'PB' not in col_str and idx_almacen == -1: idx_almacen = j
                        if ('LIBRE' in col_str or 'SALDO' in col_str) and 'VALOR' not in col_str and idx_saldo == -1: idx_saldo = j
                        
                sap_dict_pista_math = {}
                datos_extraidos_sap = []

                # ⚡ 1:1 EXTRACCIÓN EXACTA POR FILA/POSICIÓN DE SAP (Visual) y Agrupación (Matemática)
                for _, fila_sap in match_ped.iterrows():
                    col_mat = [c for c in fila_sap.index if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper() or 'COD' in str(c).upper()]
                    if not col_mat: continue
                    texto_material = str(fila_sap[col_mat[0]]).strip()
                    if "459" in texto_material or "429" in texto_material: continue

                    cod_item = texto_material.split('.')[0].lstrip('0')
                    
                    col_cant_real = [c for c in fila_sap.index if any(x in str(c).upper() for x in ['CANT', 'HECT', 'DOSIS', 'CANTIDAD'])]
                    cant_total = extraer_numero(fila_sap[col_cant_real[0]]) if col_cant_real else 0.0
                    dosis_pista = cant_total / ha_dosis_final if ha_dosis_final > 0 else 0.0

                    nombre_p = f"Item {cod_item}"
                    if not df_sab.empty:
                        df_sab_col0_clean = df_sab.iloc[:, 0].astype(str).str.split('.').str[0].str.strip().str.upper().str.lstrip('0')
                        match_sabana = df_sab[df_sab_col0_clean == cod_item]
                        if not match_sabana.empty:
                            col_nombre_sab = [c for c in match_sabana.columns if 'TEXTO' in str(c).upper() or 'DESC' in str(c).upper()]
                            if col_nombre_sab: nombre_p = str(match_sabana.iloc[0][col_nombre_sab[0]]).upper()

                    nombre_limpio = nombre_p.split('*')[0].strip().replace(" ", "")
                    # SUMA MATEMÁTICA INTERNA (Para el Motor IA)
                    sap_dict_pista_math[nombre_limpio] = sap_dict_pista_math.get(nombre_limpio, 0.0) + dosis_pista
                    # GUARDADO INDIVIDUAL (Para la Pantalla)
                    datos_extraidos_sap.append({"cod": cod_item, "nombre": nombre_p, "nombre_limpio": nombre_limpio, "cant_total": cant_total})

                # 💥 EJECUCIÓN DEL MOTOR IA CON CEREBRO CRUDO
                coctel_ganador, dosis_oficiales_coctel = emparejar_coctel_ia(sap_dict_pista_math, coctel_piloto_base)
                
                if coctel_ganador == "SIN COINCIDENCIA":
                    st.error(f"🤖 **MOTOR IA MAESTRO (Guillotina):** Cóctel Oficial Determinado: **SIN COINCIDENCIA**")
                else:
                    st.success(f"🤖 **MOTOR IA MAESTRO (Guillotina):** Cóctel Oficial Determinado: **{coctel_ganador}**")
                
                if not df_sab.empty:
                    df_sab_col0_clean = df_sab.iloc[:, 0].astype(str).str.split('.').str[0].str.strip().str.upper().str.lstrip('0')
                else: df_sab_col0_clean = pd.Series(dtype=str)

                matriz_datos = []
                for item_data in datos_extraidos_sap:
                    cod_item = str(item_data['cod']).strip().upper().lstrip('0')
                    nombre_p, nombre_limpio, cant_linea_sap = item_data['nombre'], item_data['nombre_limpio'], item_data['cant_total']
                    costo_unit, lote_sap, saldo_sap = 0.0, "SIN LOTE EN PISTA", 0.0

                    if not df_sab.empty:
                        match_sabana_global = df_sab[df_sab_col0_clean == cod_item]
                        if not match_sabana_global.empty:
                            if idx_almacen != -1: match_pista_precio = match_sabana_global[match_sabana_global.iloc[:, idx_almacen].astype(str).str.strip().str.upper().str.contains(str(pista_sel).strip().upper(), na=False)]
                            else: match_pista_precio = match_sabana_global
                            fila_precio = match_pista_precio.iloc[0] if not match_pista_precio.empty else match_sabana_global.iloc[0]
                            if idx_precio != -1: costo_unit = extraer_numero(fila_precio.iloc[idx_precio])
                            if costo_unit == 0.0:
                                col_v = [c for c in fila_precio.index if 'VALOR' in str(c).upper() and 'LIBRE' in str(c).upper()]
                                col_c = [c for c in fila_precio.index if 'LIBRE' in str(c).upper() and 'VALOR' not in str(c).upper()]
                                if col_v and col_c:
                                    v_t, c_t = extraer_numero(fila_precio[col_v[0]]), extraer_numero(fila_precio[col_c[0]])
                                    if c_t > 0: costo_unit = v_t / c_t

                            if idx_almacen != -1: match_pista = match_sabana_global[match_sabana_global.iloc[:, idx_almacen].astype(str).str.strip().str.upper().str.contains(str(pista_sel).strip().upper(), na=False)] 
                            else: match_pista = match_sabana_global

                            if not match_pista.empty:
                                fila_final = match_pista.iloc[0]
                                if idx_lote != -1: lote_sap = str(fila_final.iloc[idx_lote])
                                if idx_saldo != -1: saldo_sap = extraer_numero(fila_final.iloc[idx_saldo])

                    try:
                        if not df_cfg.empty:
                            c_p_i, c_c_i = 8, 9
                            for i_cfg in range(min(5, len(df_cfg))):
                                r_c = df_cfg.iloc[i_cfg].astype(str).str.upper().tolist()
                                if 'PRODUCTO' in r_c and 'COSTO' in r_c: 
                                    c_p_i, c_c_i = r_c.index('PRODUCTO'), r_c.index('COSTO')
                                    break
                            mask_cfg = df_cfg.iloc[:, c_p_i].astype(str).str.upper().str.strip() == nombre_limpio
                            if not mask_cfg.any(): mask_cfg = df_cfg.iloc[:, c_p_i].astype(str).str.upper().str.strip() == nombre_p.upper().strip()
                            if mask_cfg.any():
                                precio_maestro = extraer_numero(df_cfg[mask_cfg].iloc[0, c_c_i])
                                if precio_maestro > 0: costo_unit = float(precio_maestro) 
                    except Exception: pass

                    dosis_teorica = None
                    # Si encontró el cóctel, sacamos las dosis de la receta oficial
                    if coctel_ganador != "SIN COINCIDENCIA":
                        for p_receta, d_oficial in dosis_oficiales_coctel.items():
                            if p_receta == nombre_limpio or (len(nombre_limpio) >= 4 and p_receta in nombre_limpio) or (len(p_receta) >= 4 and nombre_limpio in p_receta):
                                dosis_teorica = d_oficial
                                break

                        # Reglas especiales SÓLO si el cóctel es válido
                        if dosis_teorica is None:
                            if "ACONDICIONADOR" in nombre_limpio: 
                                dosis_teorica = 0.06 if any(x in coctel_ganador for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
                            elif "ACEITE" in nombre_limpio:
                                for char in coctel_ganador:
                                    if char.isdigit():
                                        dosis_teorica = float(char)
                                        break
                            elif "IMBIOSIL" in nombre_limpio.replace(" ", ""): 
                                dosis_teorica = 1.5 if coctel_ganador.strip().upper().startswith("IN") else 1.0

                    # 💥 REGLA DE ORO: Cero adulteración. Si no hay dosis teórica oficial (porque es un intruso o porque no hay cóctel), es 0.0.
                    if dosis_teorica is None:
                        dosis_teorica = 0.0

                    dosis_ideal_pura = round(dosis_teorica * ha_dosis_final, 3)

                    matriz_datos.append({
                        "A: Producto": nombre_p, 
                        "B: Dosis Oficial/Ha": round(dosis_teorica, 3), 
                        "C: Extra %": 0.0, 
                        "D: Dosis Ideal (Total)": dosis_ideal_pura, 
                        "E: Costo Unit (+Margen)": round(costo_unit * mult_material, 0),
                        "G: Lotes (SAP)": lote_sap, 
                        "H: Saldo Real SAP": round(saldo_sap, 3), 
                        "I: Sugerido SAP (Total)": round(cant_linea_sap, 3)
                    })
                
                df_matriz = pd.DataFrame(matriz_datos)
                
                if not df_matriz.empty:
                    # 💥 REGLA DE ORO: SUMA ACUMULADA POR PRODUCTO PARA EVALUAR RECARGO REAL
                    df_matriz["TOTAL_PROD_SAP"] = df_matriz.groupby("A: Producto")["I: Sugerido SAP (Total)"].transform("sum")

                    # 💥 CÁLCULO EXACTO DE % EXTRA
                    for idx_m, r_m in df_matriz.iterrows():
                        b_ideal_pura = r_m["D: Dosis Ideal (Total)"]
                        tot_p_sap = r_m["TOTAL_PROD_SAP"]

                        if b_ideal_pura > 0:
                            extra_pct = ((tot_p_sap / b_ideal_pura) - 1.0) * 100.0
                            df_matriz.at[idx_m, "C: Extra %"] = round(extra_pct, 3)
                        else:
                            df_matriz.at[idx_m, "C: Extra %"] = 0.0

                    def calcular_semaforo_misiones(row):
                        base_pura = float(row["D: Dosis Ideal (Total)"])
                        total_producto = float(row.get("TOTAL_PROD_SAP", row["I: Sugerido SAP (Total)"]))
                        extra_pct = float(row.get("C: Extra %", 0.0))
                        diferencia = total_producto - base_pura
                        
                        if base_pura == 0.0 and total_producto > 0: return "⚠️ INTRUSO (SIN DOSIS OFICIAL)"
                        elif total_producto < (base_pura - 0.05): return "🔴 PELIGRO: SUB-DOSIS (---)"
                        elif extra_pct > 0.01 or diferencia > 0.05: return f"🔵 REC. TÉCNICA (+{extra_pct:.1f}%)"
                        else: return "🟢 ÓPTIMO"

                    df_matriz["📊 Ajuste de Campo"] = df_matriz.apply(calcular_semaforo_misiones, axis=1)
                    
                    columnas_ordenadas = [
                        "A: Producto", "B: Dosis Oficial/Ha", "C: Extra %", 
                        "D: Dosis Ideal (Total)", "I: Sugerido SAP (Total)", "📊 Ajuste de Campo",
                        "E: Costo Unit (+Margen)", "G: Lotes (SAP)", "H: Saldo Real SAP", "TOTAL_PROD_SAP"
                    ]
                    df_matriz = df_matriz[columnas_ordenadas]
                    
                    def estilizar_dosis_ideal(row):
                        estilos = [''] * len(row)
                        try:
                            idx_sistema = row.index.get_loc("D: Dosis Ideal (Total)")
                            idx_sap = row.index.get_loc("I: Sugerido SAP (Total)")
                            base_pura = float(row["D: Dosis Ideal (Total)"])
                            total_producto = float(row.get("TOTAL_PROD_SAP", row["I: Sugerido SAP (Total)"]))
                            extra_pct = float(row.get("C: Extra %", 0.0))
                            diferencia_real = total_producto - base_pura
                            
                            if base_pura == 0.0 and total_producto > 0: color = 'background-color: #ffeeba; color: #856404; font-weight: bold;'
                            elif diferencia_real < -0.05: color = 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                            elif extra_pct > 0.01 or diferencia_real > 0.05: color = 'background-color: #fff3cd; color: #856404; font-weight: bold;'
                            else: color = 'background-color: #d4edda; color: #155724; font-weight: bold;'
                                
                            estilos[idx_sistema] = color; estilos[idx_sap] = color
                        except Exception: pass
                        return estilos

                    df_estilizado = df_matriz.drop(columns=["TOTAL_PROD_SAP"]).style.apply(estilizar_dosis_ideal, axis=1)

                    edited_df = st.data_editor(
                        df_estilizado,
                        key=llave_editor_casilla,
                        column_config={
                            "B: Dosis Oficial/Ha": st.column_config.NumberColumn("Dosis Oficial/Ha", min_value=0.000, format="%.3f"),
                            "C: Extra %" : st.column_config.NumberColumn("Extra %", format="%.3f"),
                            "D: Dosis Ideal (Total)": st.column_config.NumberColumn("Dosis Ideal", format="%.3f"),
                            "I: Sugerido SAP (Total)": st.column_config.NumberColumn("Sugerido SAP (Total)", format="%.3f"),
                            "📊 Ajuste de Campo": st.column_config.TextColumn("📊 Ajuste de Campo"),
                            "E: Costo Unit (+Margen)": st.column_config.NumberColumn("Costo Unit (COP)", format="%.0f"),
                            "H: Saldo Real SAP": st.column_config.NumberColumn("Saldo SAP", format="%.3f"),
                        },
                        disabled=["A: Producto", "D: Dosis Ideal (Total)", "E: Costo Unit (+Margen)", "G: Lotes (SAP)", "H: Saldo Real SAP", "I: Sugerido SAP (Total)", "📊 Ajuste de Campo"],
                        use_container_width=True, hide_index=True
                    )

                    st.write("")
                    st.markdown("##### 📋 Copia Rápida para SAP (Costo Unitario)")
                    st.code("\n".join(df_matriz['E: Costo Unit (+Margen)'].fillna(0).astype(int).astype(str).tolist()), language="text")
            else:
                st.warning("🚨 No se encontró un pedido válido para la matriz de químicos.")

            from decimal import Decimal, ROUND_HALF_UP
            def sap_round(n): 
                n_limpio = round(float(n), 4)
                return int(Decimal(str(n_limpio)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

            if 'df_matriz' in locals() and df_matriz is not None and not df_matriz.empty:
                costo_mezcla_total = (df_matriz["I: Sugerido SAP (Total)"] * df_matriz["E: Costo Unit (+Margen)"]).apply(sap_round).sum()
            else: costo_mezcla_total = 0

            unitario_st = sap_round(d_ciclo_factura * tarifa_serv_tec_base)
            unitario_vuelo = sap_round(costo_total_vuegos / total_ha_cobro_escuadron) if total_ha_cobro_escuadron > 0 else 0
            subtotal_st_finca = sap_round(unitario_st * ha_dosis_final)
            subtotal_vuelo_finca = sap_round(unitario_vuelo * ha_dosis_final)
            gran_total = costo_mezcla_total + subtotal_vuelo_finca + subtotal_st_finca
            costo_por_ha = sap_round(gran_total / ha_dosis_final) if ha_dosis_final > 0 else 0

            precio_columna_ref = dict_aviones.get(escuadron_aviones.iloc[0]['Avión'], 0) if (not mision_solo_dron and not escuadron_aviones.empty) else 0
            precio_dron_ref = dict_drones.get(escuadron_drones.iloc[0]['Drone'], 0) if (not escuadron_drones.empty and pd.notna(escuadron_drones.iloc[0]['Drone'])) else 0

            st.write("")
            st.markdown("### 💰 Liquidación Final (Bóveda SAP)")
            m1, m2, m3, m4, m5 = st.columns(5)
            
            def mini_metric(i, t, v): return f"<div style='background-color:#ffffff; padding:12px; border-radius:8px; border: 2px solid #0d1b2a; border-left:5px solid #d4af37; box-shadow: 0 2px 4px rgba(0,0,0,0.06);'><p style='margin:0; font-size:11px; font-weight:800; color:#0d1b2a; text-transform:uppercase;'>{i} {t}</p><p style='margin:0; font-size:15px; font-weight:900; color:#1a365d;'>{v}</p></div>"
            
            with m1: st.markdown(mini_metric("🚜", "Hectáreas", f"{ha_dosis_final:.2f} Ha"), unsafe_allow_html=True)
            with m2: 
                ha_av_r = float(escuadron_aviones['Hectáreas'].sum()) if (not mision_solo_dron and not escuadron_aviones.empty) else 0
                es_dr_dom = mision_solo_dron or (ha_av_r == 0 and precio_dron_ref > 0)
                pista_display_text = "PARCELA < 20HA" if interciclo_menor_20 else tipo_de_tope_finca
                st.markdown(mini_metric("¼️", "Pista", pista_display_text if not es_dr_dom else "DRON"), unsafe_allow_html=True)
                st.markdown("<div style='margin-top:5px;'></div>", unsafe_allow_html=True)
                st.markdown(mini_metric("🚧", "Valor Tope", f"$ {fmt_sap(precio_dron_ref)}" if es_dr_dom else ("Sin Tope" if val_tope in [0, 999999] else f"$ {fmt_sap(val_tope)}")), unsafe_allow_html=True)
            with m3: st.markdown(mini_metric("👨‍🔬", "Tarifa ST", f"$ {fmt_sap(tarifa_serv_tec_base)}"), unsafe_allow_html=True)
            with m4: st.markdown(mini_metric("✈️", "Mult.", f"x {multi_aviones_final}"), unsafe_allow_html=True)
            with m5: 
                st.markdown(mini_metric("⏱️", "Precio Hora", f"$ {fmt_sap(precio_columna_ref)}"), unsafe_allow_html=True)
                st.markdown("<div style='margin-top:5px;'></div>", unsafe_allow_html=True)
                st.markdown(mini_metric("🚁", "Tarifa Dron", f"$ {fmt_sap(precio_dron_ref)}"), unsafe_allow_html=True)
            
            st.write("")
            c_sap1, c_sap2, c_sap3, c_sap4 = st.columns(4)
            with c_sap1: st.caption("👨‍🔬 UNITARIO ST (459)"); st.code(fmt_sap(unitario_st), language="text")
            with c_sap2: st.caption("✈️ UNITARIO Vuelo (429)"); st.code(fmt_sap(unitario_vuelo), language="text")
            with c_sap3: st.caption("🧪 TOTAL Mezcla"); st.code(fmt_sap(costo_mezcla_total), language="text")
            with c_sap4: st.markdown(f"<div style='background-color:#0d1b2a; padding:10px; border-radius:5px; border:2px solid #d4af37; text-align:center;'><p style='margin:0; color:#d4af37; font-size:12px; font-weight:bold;'>💰 COSTO x HA (Final)</p><h4 style='margin:0; color:white;'>$ {fmt_sap(costo_por_ha)}</h4></div>", unsafe_allow_html=True)

            html_totales = f"""
            <div style="display: flex; flex-wrap: wrap; gap: 15px; margin-top: 20px; margin-bottom: 20px;">
                <div style="flex: 1; min-width: 150px; background-color: #ffffff; padding: 15px; border-radius: 8px; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; box-shadow: 0 4px 6px rgba(0,0,0,0.08);">
                    <p style="margin:0; font-size: 12px; color: #6c757d; font-weight: bold; text-transform: uppercase;">👨‍🔬 Subtotal ST (459)</p>
                    <h3 style="margin:0; color: #0d1b2a; font-weight: 900; user-select: all;">$ {fmt_sap(subtotal_st_finca)}</h3>
                </div>
                <div style="flex: 1; min-width: 150px; background-color: #ffffff; padding: 15px; border-radius: 8px; border: 2px solid #0d1b2a; border-left: 6px solid #1a365d; box-shadow: 0 4px 6px rgba(0,0,0,0.08); margin-bottom: 20px;">
                    <p style="margin:0; font-size: 12px; color: #6c757d; font-weight: bold; text-transform: uppercase;">✈️ Subtotal Vuelo (429)</p>
                    <h3 style="margin:0; color: #0d1b2a; font-weight: 900; user-select: all;">$ {fmt_sap(subtotal_vuelo_finca)}</h3>
                </div>
                <div style="flex: 1.5; min-width: 200px; background-color: #0d1b2a; padding: 15px; border-radius: 8px; border: 3px solid #d4af37; box-shadow: 0 4px 12px rgba(0,0,0,0.2); text-align: center;">
                    <p style="margin:0; font-size: 13px; color: #d4af37; font-weight: bold; text-transform: uppercase;">🔥 TOTAL OPERACIÓN</p>
                    <h2 style="margin:0; color: white; font-weight: 900; user-select: all;">$ {gran_total:,.0f}</h2>
                </div>
            </div>
            """.replace(",", ".")
            st.markdown(html_totales, unsafe_allow_html=True)
            
            st.caption("📋 **COPIA RÁPIDA (Clic en el ícono 📋 de cada cajita)**")
            cc1, cc2, cc3, cc4, cc5, cc6 = st.columns(6)
            with cc1: st.write("👨‍🔬 Serv. Tec"); st.code(fmt_sap(subtotal_st_finca), language="text")
            with cc2: st.write("✈️ Vuelo"); st.code(fmt_sap(subtotal_vuelo_finca), language="text")
            with cc3: st.write("🧪 Mezcla"); st.code(fmt_sap(costo_mezcla_total), language="text")
            with cc4: st.write("⚠️ Recargo"); st.code(fmt_sap(0), language="text")
            with cc5: st.write("💰 Costo x Ha"); st.code(fmt_sap(costo_por_ha), language="text")
            with cc6: st.write("🔥 TOTAL"); st.code(fmt_sap(gran_total), language="text")

            st.markdown("---")
            st.markdown("### 🛰️ Coordenadas de Lanzamiento Final")
            c_p1, c_p2 = st.columns(2)
            pista_manual = c_p1.selectbox("📍 Confirmar Pista de Operación:", ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"], index=["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"].index(pista_sel), key=f"confirmador_final_{pista_sel}_{vuelo_ref}")
            c_p2.info(f"🚀 Misión: {('DRONE' if mision_solo_dron else 'AVION')} | 📋 Referencia: {vuelo_ref}")
            
            st.markdown("""
                <a href="#inicio_modulo" target="_self" style="
                    display: block; width: 100%; text-align: center; 
                    background-color: #0d1b2a; color: #d4af37; border: 1px solid #d4af37; 
                    padding: 12px; border-radius: 8px; text-decoration: none; font-weight: bold;
                    box-shadow: 0px 4px 6px rgba(0,0,0,0.3); margin-bottom: 20px; font-size: 16px;
                ">
                    ⬆️ VOLVER AL INICIO DEL MÓDULO ⬆️
                </a>
            """, unsafe_allow_html=True)

            if st.button("💾 DETONAR FACTURA Y GUARDAR EN BÓVEDA", type="primary", use_container_width=True):
                with st.spinner("🚀 Inyectando datos con Precisión de Francotirador a Velocidad Luz..."):
                    try:
                        gc_save = obtener_cliente_gspread_unificado()
                        if not gc_save: st.error("🚨 Error crítico: No se pudo conectar al llavero de Google para guardar."); st.stop()

                        boveda = gc_save.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                        hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023")
                        hoja_maestra = boveda.worksheet("TABLA 1")
                        hoja_memoria = boveda.worksheet("MEMORIA")

                        fecha_str = fecha_operacion.strftime("%d/%m/%Y")
                        dia_sem = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"][fecha_operacion.weekday()]
                        num_sem = fecha_operacion.isocalendar()[1]
                        os_virtual = f"VIRT-{finca_limpia[:3]}-{datetime.now().strftime('%H%M')}"
                        
                        bloque_f, sector_f, ha_bruta_f = "", "", ""
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_limpia.upper().strip()]
                        if not match_f.empty: sector_f, ha_bruta_f, bloque_f = match_f.iloc[0, 1], match_f.iloc[0, 2], match_f.iloc[0, 3]

                        ha_f = float(ha_dosis_final)
                        if pd.isna(ha_bruta_f) or str(ha_bruta_f).strip() == "": ha_bruta_f = ha_f
                        
                        h_total_v = 0.0 if total_ha_cobro_escuadron == 0 else (ha_f / 10 if mision_solo_dron else ((ha_f / total_ha_cobro_escuadron) * horometro_final_avion))
                        vol_total_gln, rend_min = ha_f * 6, h_total_v * 60
                        
                        if total_ha_cobro_escuadron == 0:
                            piloto_f, hk_f, tipo_mision_str = "CLIENTE (DRON PROPIO)", "DRON PROPIO", "DRONE PROPIO"
                        else:
                            piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                            tipo_mision_str = 'DRONE' if mision_solo_dron else 'AVION'
                            if mision_solo_dron:
                                if not escuadron_drones.empty:
                                    dr_name_str = str(escuadron_drones.iloc[0].get('Drone', '')).upper()
                                    hk_f = "DR51" if "DATAROT" in dr_name_str else ("DR52" if "GENESYS" in dr_name_str else "DRONE_GEN")
                                else: hk_f = "DRONE_GEN"
                            else:
                                hk_f = str(escuadron_aviones.iloc[0].get('Avión', 'AVION_REG')).upper() if not escuadron_aviones.empty else "AVION_REG"
                        
                        tarifa_vuelo_neta_ha = float(costo_neto_vuelo_total / total_ha_cobro_escuadron) if total_ha_cobro_escuadron > 0 else 0.0
                        total_pago_avion_neto = (tarifa_vuelo_neta_ha + float(recargo_final)) * ha_f
                        
                        row_azul = [""] * 34
                        row_azul[0], row_azul[1], row_azul[2], row_azul[3], row_azul[4], row_azul[5] = os_virtual, bloque_f, finca_limpia, sector_f, ha_bruta_f, ha_f
                        row_azul[6], row_azul[7], row_azul[8], row_azul[9], row_azul[10] = coctel_ganador, fecha_str, dia_sem, num_sem, round(h_total_v, 2)
                        row_azul[11], row_azul[12], row_azul[13], row_azul[14], row_azul[15] = 6, round(vol_total_gln, 2), round(h_total_v, 2), round(rend_min, 2), piloto_f
                        row_azul[16], row_azul[17], row_azul[18], row_azul[19], row_azul[20] = hk_f, tipo_mision_str, float(gran_total), round(tarifa_vuelo_neta_ha, 2), round(float(recargo_final), 2)
                        row_azul[21], row_azul[23], row_azul[28], row_azul[29], row_azul[32], row_azul[33] = float(gran_total), pista_manual, float(gran_total), round(total_pago_avion_neto, 2), tipo_productor, "GÉNESIS_V2_PRO"
                        
                        fila_apoyo = ["", finca_limpia, ha_f, float(costo_por_ha), float(gran_total), fecha_str, "", "", coctel_ganador, "", pista_manual, "", "", tipo_mision_str, ""]
                        
                        col_azul, col_apoyo = hoja_maestra.col_values(1), hoja_apoyo.col_values(2)
                        datos_memoria = hoja_memoria.get_all_values()

                        f_azul = next((i+2 for i in range(len(col_azul)-1, -1, -1) if str(col_azul[i]).strip() != ""), 1)
                        f_apoyo = next((i+2 for i in range(len(col_apoyo)-1, -1, -1) if str(col_apoyo[i]).strip() != ""), 1)
                        fila_apoyo[0] = f_apoyo - 3
                        
                        if f_azul > hoja_maestra.row_count: hoja_maestra.add_rows(10)
                        if f_apoyo > hoja_apoyo.row_count: hoja_apoyo.add_rows(10)
                        
                        set_existentes = {f"{str(r[0]).strip()}|{str(r[9]).strip().upper()}|{str(r[3]).strip().upper()}" for r in datos_memoria[1:] if len(r) >= 10}
                        bodega_f = "BODEGA PRINCIPAL DRON" if mision_solo_dron or total_ha_cobro_escuadron == 0 else "BODEGA PRINCIPAL AVIÓN"
                        filas_memoria = []
                        
                        if not edited_df.empty:
                            for idx, row_m in edited_df.iterrows():
                                nombre_prod = str(row_m.get("A: Producto", "")).strip().upper()
                                if nombre_prod not in ["", "NAN"]:
                                    if f"{fecha_str}|{finca_limpia}|{nombre_prod}" not in set_existentes:
                                        fila_m = [fecha_str, coctel_ganador, str(pista_manual).split("-")[0].strip()[:4], nombre_prod, str(row_m.get("G: Lotes (SAP)", "S/N")), float(row_m.get("D: Dosis Ideal (Total)", 0)), bodega_f, "", "X", finca_limpia]
                                        filas_memoria.append(fila_m)
                        
                        def limpiar_json(val): return "" if pd.isna(val) or (isinstance(val, float) and math.isnan(val)) else val
                        row_azul = [limpiar_json(x) for x in row_azul]; fila_apoyo = [limpiar_json(x) for x in fila_apoyo]; filas_memoria = [[limpiar_json(x) for x in fila] for fila in filas_memoria]

                        hoja_maestra.update(range_name=f"A{f_azul}", values=[row_azul], value_input_option='USER_ENTERED')
                        hoja_apoyo.update(range_name=f"A{f_apoyo}", values=[fila_apoyo], value_input_option='USER_ENTERED')
                        if filas_memoria: hoja_memoria.append_rows(filas_memoria, value_input_option='USER_ENTERED')

                        if 'supabase' in st.session_state:
                            try:
                                supabase_client = st.session_state['supabase']
                                payload_orden = {"os_virtual": str(os_virtual), "finca": str(finca_limpia), "hectareas": float(ha_f), "coctel": str(coctel_ganador), "fecha": str(fecha_str), "total_operacion": float(gran_total), "pista": str(pista_manual), "tipo_productor": str(tipo_productor)}
                                supabase_client.table("facturas_detonadas").insert(payload_orden).execute()
                            except Exception: pass

                        st.balloons(); st.success(f"✅ IMPACTO TOTAL CONFIRMADO. Guardado en fila {f_azul} de Drive y replicado en la bóveda relacional.")
                        st.toast("💾 Memoria Sincronizada con éxito.", icon="⚔️")
                        if 'memoria_excel' in st.session_state: del st.session_state['memoria_excel']
                    except Exception as e_save: st.error(f"🚨 Falla en el Guardado: {e_save}")

if __name__ == "__main__":
    pass
