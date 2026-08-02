import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta, date
import re
import io
from difflib import get_close_matches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --- 🔌 CONEXIÓN Y TIEMPO ---
def obtener_hora_colombia():
    return datetime.utcnow() + timedelta(hours=-5)

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception: return None

@st.cache_data(show_spinner=False, ttl=300)
def obtener_datos_bovedas():
    gc = inicializar_cliente_gspread()
    if not gc: return None, None, None, None, "No hay conexión con Google Cloud"
    URL_ING = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"
    URL_TRA = "https://docs.google.com/spreadsheets/d/1JV-f8zzGuhGNlqvrSjeKYN4eBdshAN5EOkfDHMi1WIs/edit"
    try:
        sh_ing = gc.open_by_url(URL_ING)
        ws_ing = sh_ing.worksheets()[0]
        datos_ing = ws_ing.get_all_values()
        try: datos_dicc = sh_ing.worksheet("DICCIONARIO").get_all_values()
        except: datos_dicc = []
        sh_tras = gc.open_by_url(URL_TRA)
        ws_tras = sh_tras.worksheets()[0] 
        datos_tras = ws_tras.get_all_values()
        titulo_tras = ws_tras.title 
        return datos_ing, datos_dicc, datos_tras, titulo_tras, None
    except Exception as e: return None, None, None, None, str(e)

# --- 🔍 RASTREADOR DE MATERIALES DINÁMICO (💥 CON CEBO TÁCTICO) ---
@st.cache_data(show_spinner=False, ttl=0) # TTL 0 PARA FORZAR EL CEBO A REPETIRSE SIEMPRE
def extraer_mapeo_materiales():
    gc = inicializar_cliente_gspread()
    if not gc: return {}, "FALLO: No hay conexión con Google."
    try:
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHalOnmFUJQYFggARP4/edit")
        ws = sh.worksheet("Plantilla")
        datos = ws.get_all_values()
        mapeo = {}
        cebo = []
        
        if not datos: return {}, "FALLO: La pestaña Plantilla está vacía."
        
        # Encontrar en qué fila están los títulos realmente
        idx_head = -1
        for i in range(min(15, len(datos))):
            fila_up = [str(x).strip().upper() for x in datos[i]]
            if "MATERIAL" in fila_up:
                idx_head = i
                break
                
        if idx_head == -1:
            return {}, f"FALLO: No encontré la palabra 'MATERIAL' en las primeras 15 filas. Filas iniciales: {datos[:3]}"
            
        encabezados = [str(x).strip().upper() for x in datos[idx_head]]
        cebo.append(f"📍 Títulos en fila {idx_head + 1}.")
        
        idx_mat = encabezados.index("MATERIAL")
        
        # Buscar Columna J o K
        idx_desc_j = encabezados.index("DESCRIPCIÓN DEL MATERIAL") if "DESCRIPCIÓN DEL MATERIAL" in encabezados else -1
        idx_desc_k = encabezados.index("DESCRIPCIÓN ÚNICA") if "DESCRIPCIÓN ÚNICA" in encabezados else -1
        
        cebo.append(f"📊 Índices - MAT: {idx_mat}, DESC_J: {idx_desc_j}, DESC_K: {idx_desc_k}")

        for row in datos[idx_head+1:]:
            mat = str(row[idx_mat]).strip() if len(row) > idx_mat else ""
            desc = ""
            
            # Intentar primero Columna K, si está vacía, intentar Columna J
            if idx_desc_k != -1 and len(row) > idx_desc_k and str(row[idx_desc_k]).strip():
                desc = str(row[idx_desc_k]).strip().upper()
            elif idx_desc_j != -1 and len(row) > idx_desc_j and str(row[idx_desc_j]).strip():
                desc = str(row[idx_desc_j]).strip().upper()
                
            if desc and mat:
                desc_clean = re.sub(r'\s+', ' ', desc).strip()
                mapeo[desc_clean] = mat
                
        cebo.append(f"✅ Total materiales memorizados: {len(mapeo)}.")
        # Extraemos unos cuantos para ver si los está leyendo bien
        muestras = list(mapeo.items())[:3]
        cebo.append(f"🔍 Muestra: {muestras}")
        
        return mapeo, " | ".join(cebo)
    except Exception as e: return {}, f"ERROR TÉCNICO: {str(e)}"

def buscar_codigo_material(producto_nombre, mapeo):
    prod_clean = str(producto_nombre).strip().upper()
    if not prod_clean or not mapeo: return "S/N"
    
    # 1. Match Exacto
    if prod_clean in mapeo: return mapeo[prod_clean]
    
    # 2. Match Parcial (Ej: "BOSCALID" entra en "BOSCALID 50 WG")
    for desc, cod in mapeo.items():
        if prod_clean in desc or desc in prod_clean: return cod
        
    # 3. Match Aproximado
    matches = get_close_matches(prod_clean, list(mapeo.keys()), n=1, cutoff=0.6)
    if matches: return mapeo[matches[0]]
    
    return "S/N"

# 💥 RADAR CRONOLÓGICO Y ANTI-FALLOS
def procesar_fecha_estricta(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() in ["none", "nan", "nat", "<na>"]: return pd.NaT
    s = str(val).strip().lower()
    if s.replace('.', '', 1).isdigit(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
    meses_es = {'enero':1, 'febrero':2, 'marzo':3, 'abril':4, 'mayo':5, 'junio':6, 'julio':7, 'agosto':8, 'septiembre':9, 'octubre':10, 'noviembre':11, 'diciembre':12}
    mes_encontrado = next((meses_es[m] for m in meses_es if m in s), None)
    if mes_encontrado:
        numeros = re.findall(r'\d+', s)
        if len(numeros) >= 2:
            n1, n2 = int(numeros[0]), int(numeros[1])
            anio = n1 if n1 > 1000 else (n2 if n2 > 1000 else (2000 + n2 if n2 < 100 else n2))
            dia = n2 if n1 > 1000 else (n1 if n2 > 1000 else n1)
            try: return pd.Timestamp(year=anio, month=mes_encontrado, day=dia)
            except: pass
    for dia_sem in ['lunes','martes','miércoles','miercoles','jueves','viernes','sábado','sabado','domingo']: s = s.replace(dia_sem, '')
    s = s.replace(',', '').replace(' de ', '/').replace('-', '/').strip()
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y'):
        try: return pd.to_datetime(s, format=fmt)
        except: pass
    try: 
        res = pd.to_datetime(s, dayfirst=True)
        return pd.NaT if pd.isna(res) else res
    except: return pd.NaT 

def formatear_numero_sap(val):
    try:
        f_val = float(str(val).replace(",", ""))
        if f_val.is_integer(): return f"{int(f_val):,}".replace(",", ".")
        val_str = f"{f_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return val_str[:-3] if val_str.endswith(",00") else val_str
    except: return str(val)

def estandarizar_pista(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() in ["none", "nan", "nat", "<na>"]: return ""
    return str(val).strip().upper().replace("ORIHUECA", "PORI").replace("DIVAS", "PDIV").replace("TEHOBROMINA", "TEHO").replace("LUCHA", "PLUC")

@st.cache_data(show_spinner=False, ttl=3600)
def extraer_catalogo_oficial_sap():
    gc = inicializar_cliente_gspread()
    if not gc: return []
    try:
        sh_config = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHalOnmFUJQYFggARP4/edit")
        ws = sh_config.worksheet("Configuración") if "Configuración" in [w.title for w in sh_config.worksheets()] else sh_config.worksheet("DD_Mesclas")
        datos = ws.get_all_values()
        if not datos: return []
        productos_oficiales = set()
        idx_prod = -1
        for i in range(min(10, len(datos))):
            if 'PRODUCTO' in [str(x).upper().strip() for x in datos[i]]:
                idx_prod = [str(x).upper().strip() for x in datos[i]].index('PRODUCTO')
                break
        if idx_prod != -1:
            for row in datos[i+1:]:
                if len(row) > idx_prod:
                    p_nombre = re.sub(r'\s+', ' ', str(row[idx_prod]).strip().upper())
                    if p_nombre and len(p_nombre) > 2 and p_nombre != "0": productos_oficiales.add(p_nombre)
        return list(productos_oficiales)
    except: return []

DICT_BASE_PRODUCTOS = {"ACEITE DICAM": "ROYAL BIOCHEM", "ACONDICIONADOR SV": "SYS TECNOLOGIES", "ADHERENTE SV": "SYS TECNOLOGIES", "BANADAK": "PLANDAK", "BANANO Y PLATANO * LT": "INVESA S.A.S.", "BANATREL SC": "YARA S.A.S.", "BOSCALID 50 WG": "DVA COLOMBIA", "CERAQUINT SP": "CERADIS COLOMBIA", "CEROSTRESS SV * LT": "MICROFERTIZA", "COMPER SV": "ADAMA", "EPOXICONAZOLE DEL MONTE": "DEL MONTE SAS", "FENTRIUPH AGRO 88 OL": "DEL MONTE SAS", "FOSFOSTRESS SV": "MICROFERTIZA", "GLOBAFOL nf": "SYNGENTA", "IMBIOSIL O": "INBIOMA", "KURDO 250 EC": "INVESA S.A.S.", "KYVENTIQ": "CORTEVA", "LONSELOR 30 SC": "BASF QUÍMICA", "MANCOL 430 SC": "CASAGRO", "NATURAMIN WSP": "AGRIANDES DAINSA", "OPORTO": "ADAMA", "OPUS 12 EC": "BASF QUÍMICA", "POLYTHION SC": "UPL", "POWMYL SV": "SUMITOMO", "QUELAMIX": "INGEPLANT", "REFLECT": "SYNGENTA", "ROUTINE SC": "BAYER", "SEEKER": "SYNGENTA", "SICO": "SYNGENTA", "SIGANEX 60 SC": "BAYER", "SPRAYFIX": "AGRIANDES DAINSA", "THIOPRON 825 SC": "UPL", "TIMOREX PRO": "ADAMA", "XILOTROM": "AGRIFOL", "ZINTRAC x LITRO SV": "YARA S.A.S."}

def limpiar_celda_none(val):
    v = str(val).strip()
    return "" if v.lower() in ["none", "nan", "nat", "<na>", "null"] else v

# --- 🚀 EJECUCIÓN DEL MÓDULO ---
def ejecutar():
    hoy_colombia = obtener_hora_colombia().date()
    if 'form_key_m19' not in st.session_state: st.session_state['form_key_m19'] = 0
    if 'form_key_m19_traslados' not in st.session_state: st.session_state['form_key_m19_traslados'] = 0

    def limpiar_campos_operativos(): st.session_state['form_key_m19'] += 1
    def limpiar_campos_traslados(): st.session_state['form_key_m19_traslados'] += 1

    st.markdown("<div id='inicio-modulo-19'></div>", unsafe_allow_html=True)
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    .kpi-card { background-color: #0d1b2a; color: white; padding: 20px; border-radius: 10px; border-left: 6px solid #d4af37; box-shadow: 0 4px 6px rgba(0,0,0,0.2); margin-bottom: 15px; }
    .kpi-rojo { border-left-color: #dc3545; } .kpi-amarillo { border-left-color: #ffc107; } .kpi-verde { border-left-color: #28a745; } .kpi-azul { border-left-color: #2F75B5; }
    .kpi-titulo { font-weight: bold; font-size: 14px; margin-bottom: 5px; text-transform: uppercase; color: #a0aec0; }
    .kpi-valor { font-size: 28px; font-weight: 900; margin: 0; color: white; }
    div[data-testid="stExpander"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.15) !important; background-color: #ffffff !important; margin-bottom: 20px !important; }
    div[data-testid="stExpander"] summary { background-color: #0d1b2a !important; border-radius: 6px 6px 0px 0px !important; padding: 10px 15px !important; }
    div[data-testid="stExpander"] summary p { color: #d4af37 !important; font-family: 'Arial Black', sans-serif !important; font-size: 15px !important; text-transform: uppercase !important; margin: 0 !important; }
    div[data-testid="stMainBlockContainer"] label p { color: #0d1b2a !important; font-weight: 900 !important; text-transform: uppercase !important; font-size: 13px !important; }
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stDateInput"] input { border: 2px solid #0d1b2a !important; border-radius: 6px !important; color: #000000 !important; font-weight: 900 !important; background-color: #ffffff !important; }
    div[data-testid="stSelectbox"] > div:last-child { border: 2px solid #0d1b2a !important; border-radius: 6px !important; background-color: #ffffff !important; }
    div[data-testid="stSelectbox"] > div:last-child * { color: #000000 !important; font-weight: 900 !important; }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 6px 15px rgba(0,0,0,0.15) !important; background-color: #ffffff !important; }
    div[data-testid="stTabs"] button[role="tab"] { font-family: 'Arial Black', sans-serif; font-size: 14px; text-transform: uppercase; color: #0d1b2a; }
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { border-bottom-color: #d4af37; background-color: rgba(212, 175, 55, 0.1); }
    .btn-ascensor { display: block; width: 100%; text-align: center; background-color: #15283c; color: #d4af37 !important; padding: 12px; border-radius: 8px; text-decoration: none !important; font-weight: 900; border: 2px solid #d4af37; margin-bottom: 20px; box-shadow: 0px 4px 6px rgba(0,0,0,0.2); transition: all 0.3s ease; }
    .btn-ascensor:hover { background-color: #0d1b2a; box-shadow: 0px 0px 10px rgba(212, 175, 55, 0.8); }
    </style>
    """, unsafe_allow_html=True)

    URL_SHEET_INGRESOS = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"
    URL_SHEET_TRASLADOS = "https://docs.google.com/spreadsheets/d/1JV-f8zzGuhGNlqvrSjeKYN4eBdshAN5EOkfDHMi1WIs/edit"

    c_tit, c_btn1, c_btn2 = st.columns([2, 1, 1])
    c_tit.markdown("<h1 class='titulo-mod'>📦 19. Centro Logístico Unificado</h1>", unsafe_allow_html=True)
    if c_btn1.button("🔄 REFRESCAR RADARES", type="primary", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    if c_btn2.button("🧹 PURGAR DICCIONARIO", type="secondary", use_container_width=True):
        with st.spinner("Aniquilando diccionario oculto..."):
            gc = inicializar_cliente_gspread()
            if gc:
                try:
                    sh_local = gc.open_by_url(URL_SHEET_INGRESOS)
                    ws_dicc = sh_local.worksheet("DICCIONARIO")
                    sh_local.del_worksheet(ws_dicc)
                    st.cache_data.clear(); st.success("✅ ¡Fantasma destruido!"); st.rerun()
                except Exception: st.info("No había basura. La lista ya estaba limpia.")

    with st.spinner("📡 Sincronizando Bóvedas de Ingresos y Traslados (Caché Activo)..."):
        datos_ing_crudos, datos_dicc_crudos, datos_tras_crudos, titulo_ws_traslados, error_api = obtener_datos_bovedas()
        if error_api:
            st.error("🚨 **LÍMITE DE PROTECCIÓN DE GOOGLE (Error 429)**: Espera de 1 a 2 minutos y presiona 'Refrescar Radares'." if "429" in error_api else f"🚨 Error de acceso a las Bóvedas. Detalle: {error_api}")
            return

        # 🚨 CEBO 1: EXTRACCIÓN DE MATERIALES
        mapeo_materiales, info_cebo_plantilla = extraer_mapeo_materiales()
        st.error(f"🚨 CEBO 1 (PLANTILLA): {info_cebo_plantilla}")

        dict_operativo = {k.upper().strip(): v.upper().strip() for k, v in DICT_BASE_PRODUCTOS.items()}
        for row in datos_dicc_crudos[1:]:
            if len(row) >= 2 and str(row[0]).strip():
                p_nube = re.sub(r'\s+', ' ', str(row[0]).strip().upper())
                p_prov = str(row[1]).strip().upper()
                if len(p_nube) > 3: dict_operativo[p_nube] = p_prov
                
        productos_precio_sap = extraer_catalogo_oficial_sap()
        for prod_sap in productos_precio_sap:
            if prod_sap.strip() not in dict_operativo: dict_operativo[prod_sap.strip()] = "" 
        lista_autorizada = set([re.sub(r'\s+', ' ', str(k).strip().upper()) for k in dict_operativo.keys()])
        for p in productos_precio_sap: lista_autorizada.add(re.sub(r'\s+', ' ', str(p).strip().upper()))
            
        cache_mapeo = {}
        def estandarizar_y_marcar_inteligente(prod):
            p_clean = re.sub(r'\s+', ' ', str(prod).strip().upper())
            if p_clean in ["", "NONE", "NAN", "NAT", "<NA>"]: return ""
            if p_clean in cache_mapeo: return cache_mapeo[p_clean]
            resultado = ""
            if p_clean in lista_autorizada: resultado = p_clean
            else:
                posibles = [o for o in lista_autorizada if (p_clean in o) or (o in p_clean)]
                if posibles:
                    posibles.sort(key=lambda x: abs(len(x) - len(p_clean)))
                    resultado = posibles[0]
                else:
                    matches = get_close_matches(p_clean, list(lista_autorizada), n=1, cutoff=0.65)
                    resultado = matches[0] if matches else f"{p_clean} 🛑 [OBSOLETO]"
            cache_mapeo[p_clean] = resultado
            return resultado

        # ================= PROCESAMIENTO TRASLADOS =================
        df_traslados = pd.DataFrame()
        encabezados_limpios_tras = []
        if datos_tras_crudos:
            columnas_oficiales = ["CONSECUTIVO", "FECHA", "PRODUCTO", "CANTIDAD", "UNIDAD", "PISTA", "SEMANA", "OBSERVACION", "LOTE"]
            idx_head_t = next((i for i, row in enumerate(datos_tras_crudos[:15]) if "CONSECUTIVO" in [str(x).upper().strip() for x in row]), -1)
            if idx_head_t != -1:
                enc_brutos = [str(x).strip().upper() for x in datos_tras_crudos[idx_head_t]]
                for j, h in enumerate(enc_brutos):
                    if "OBSERVAC" in h or ("COLUMNA1" in h and "OBSERVACION" not in encabezados_limpios_tras): h = "OBSERVACION"
                    elif "LOTE" in h: h = "LOTE"
                    elif h == "": h = f"COL_VACIA_{j}"
                    encabezados_limpios_tras.append(h)
                
                if "OBSERVACION" not in encabezados_limpios_tras: encabezados_limpios_tras.append("OBSERVACION")
                if "LOTE" not in encabezados_limpios_tras: encabezados_limpios_tras.append("LOTE")
                
                max_cols_t = len(encabezados_limpios_tras)
                datos_recortados = []
                for row in datos_tras_crudos[idx_head_t+1:]:
                    pad_row = row + [""] * (max_cols_t - len(row))
                    pad_row = ["" if str(x).strip().lower() in ["none", "nan", "null", "nat"] else str(x).strip() for x in pad_row[:max_cols_t]]
                    datos_recortados.append(pad_row)
                    
                df_t = pd.DataFrame(datos_recortados, columns=encabezados_limpios_tras)
                df_t = df_t.loc[:, ~df_t.columns.str.startswith('COL_VACIA_')]
                df_t = df_t.loc[:, ~df_t.columns.duplicated()]
                df_t['FILA_EXCEL'] = range(idx_head_t + 2, len(df_t) + idx_head_t + 2)
                cols_presentes = [c for c in columnas_oficiales if c in df_t.columns] + ['FILA_EXCEL']
                df_traslados = df_t[cols_presentes]
                if "PRODUCTO" in df_traslados.columns: df_traslados["PRODUCTO"] = df_traslados["PRODUCTO"].apply(estandarizar_y_marcar_inteligente)
                if "PISTA" in df_traslados.columns: df_traslados["PISTA"] = df_traslados["PISTA"].apply(estandarizar_pista)

        # ================= PROCESAMIENTO INGRESOS =================
        df_ingresos = pd.DataFrame()
        encabezados_limpios_ing = []
        if datos_ing_crudos:
            idx_head_i = next((i for i, row in enumerate(datos_ing_crudos[:5]) if "ESTADO / OBSERVACIÓN" in [str(x).upper().strip() for x in row] or "PRODUCTO" in [str(x).upper().strip() for x in row]), 0)
            enc_brutos_i = [str(x).strip().upper() for x in datos_ing_crudos[idx_head_i]]
            for j, h in enumerate(enc_brutos_i):
                if h == "": h = f"COL_VACIA_{j}"
                encabezados_limpios_ing.append(h)
                
            max_cols_ing = len(encabezados_limpios_ing)
            datos_ing_recortados = []
            for row in datos_ing_crudos[idx_head_i+1:]:
                pad_row = row + [""] * (max_cols_ing - len(row))
                pad_row = ["" if str(x).strip().lower() in ["none", "nan", "null", "nat"] else str(x).strip() for x in pad_row[:max_cols_ing]]
                datos_ing_recortados.append(pad_row)
                
            df_i = pd.DataFrame(datos_ing_recortados, columns=encabezados_limpios_ing)
            df_i = df_i.loc[:, ~df_i.columns.str.startswith('COL_VACIA_')]
            df_i = df_i.loc[:, ~df_i.columns.duplicated()]
            df_i['FILA_EXCEL'] = range(idx_head_i + 2, len(df_i) + idx_head_i + 2)
            if "PRODUCTO" in df_i.columns: df_i["PRODUCTO"] = df_i["PRODUCTO"].apply(estandarizar_y_marcar_inteligente)
            if "PISTA" in df_i.columns: df_i["PISTA"] = df_i["PISTA"].apply(estandarizar_pista)
            df_ingresos = df_i

    tab_ingresos, tab_traslados = st.tabs(["📥 1. INGRESOS (COMPRAS/PROVEEDOR)", "🚚 2. MOVIMIENTOS INTERNOS (TRASLADOS)"])

    # ========================================================================
    # 📥 PESTAÑA 1: INGRESOS 
    # ========================================================================
    with tab_ingresos:
        st.markdown(f"<a href='{URL_SHEET_INGRESOS}' target='_blank' class='btn-ascensor' style='background-color:#1e4620; border-color:#2e7d32; color:#ffffff !important;'>👁️ VER BASE DE INGRESOS EN GOOGLE SHEETS</a>", unsafe_allow_html=True)
        st.write("Panel táctico de ingresos. Registra lotes de proveedores externos cruzando información oficial con la Base de Precios SAP.")

        if df_ingresos.empty:
            st.warning("La base de datos de ingresos parece estar vacía.")
        else:
            df = df_ingresos
            COL_ESTADO = "ESTADO / OBSERVACIÓN"
            col_producto = next((c for c in df.columns if "PRODUCTO" in c), None)
            col_pista = next((c for c in df.columns if "PISTA" in c), None)
            col_fecha_ingreso = next((c for c in df.columns if "FECHA DE INGRESO" in c), None)
            col_fv = next((c for c in df.columns if c in ["F/V", "FECHA VENCIMIENTO", "VENCIMIENTO"]), None)

            if COL_ESTADO not in df.columns: st.error(f"🚨 FALTA COLUMNA TÁCTICA: No se encontró la columna **{COL_ESTADO}**.")
            else:
                idx_col_estado = encabezados_limpios_ing.index(COL_ESTADO) + 1 
                st.markdown("### 📡 Radares de Vencimiento")
                limite_90_dias = pd.to_datetime(hoy_colombia) + pd.to_timedelta(90, unit='D')
                hoy_ts = pd.to_datetime(hoy_colombia)
                
                lotes_vencidos, lotes_riesgo = 0, 0
                if col_fv:
                    df['FECHA_VENC_DT'] = df[col_fv].apply(procesar_fecha_estricta)
                    df_activos = df[~df[COL_ESTADO].str.contains("ANULADO|ELIMINAR", na=False, case=False)]
                    lotes_vencidos = df_activos[df_activos['FECHA_VENC_DT'] < hoy_ts].shape[0]
                    lotes_riesgo = df_activos[(df_activos['FECHA_VENC_DT'] >= hoy_ts) & (df_activos['FECHA_VENC_DT'] <= limite_90_dias)].shape[0]

                k1, k2, k3 = st.columns(3)
                k1.markdown(f"<div class='kpi-card kpi-verde'><div class='kpi-titulo'>📦 Ingresos Registrados</div><p class='kpi-valor'>{len(df)}</p></div>", unsafe_allow_html=True)
                k2.markdown(f"<div class='kpi-card kpi-rojo'><div class='kpi-titulo'>🚨 Lotes Vencidos</div><p class='kpi-valor'>{lotes_vencidos}</p></div>", unsafe_allow_html=True)
                k3.markdown(f"<div class='kpi-card kpi-amarillo'><div class='kpi-titulo'>⚠️ Por Vencer (90 Días)</div><p class='kpi-valor'>{lotes_riesgo}</p></div>", unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("""<a href="#seccion-auditoria" class="btn-ascensor">👇 SALTAR DIRECTO A LA MATRIZ DE AUDITORÍA 👇</a>""", unsafe_allow_html=True)
                st.markdown("### ➕ Inyector de Nuevos Ingresos")

                with st.expander("🧪 1. IDENTIFICACIÓN DEL QUÍMICO OFICIAL", expanded=True):
                    c_tog1, c_tog2 = st.columns(2)
                    es_nuevo_producto = c_tog1.toggle("✨ Ingresar un Producto Totalmente NUEVO")
                    modificar_prov = False
                    c_prod, c_mat, c_prov = st.columns([2, 1, 2])
                    
                    if es_nuevo_producto:
                        n_prod = c_prod.text_input("🧪 Nombre del Nuevo Producto")
                        n_mat = c_mat.text_input("🔢 Cód. Material", placeholder="No aplica", disabled=True)
                        mat_item_ing = "S/N"
                        n_prov = c_prov.text_input("🏭 Nombre del Proveedor")
                        with c_tog2:
                            st.markdown("<div style='margin-top: -5px;'></div>", unsafe_allow_html=True)
                            if st.button("💾 GUARDAR EN DICCIONARIO", type="secondary", use_container_width=True):
                                prod_limpio = str(n_prod).strip().upper(); prov_limpio = str(n_prov).strip().upper()
                                if not prod_limpio or not prov_limpio: st.warning("⚠️ Escribe nombre y proveedor.")
                                else:
                                    with st.spinner("Registrando nuevo químico..."):
                                        try:
                                            gc_temp = inicializar_cliente_gspread()
                                            sh_temp = gc_temp.open_by_url(URL_SHEET_INGRESOS)
                                            try: ws_dicc = sh_temp.worksheet("DICCIONARIO")
                                            except:
                                                ws_dicc = sh_temp.add_worksheet(title="DICCIONARIO", rows="100", cols="2")
                                                ws_dicc.append_row(["PRODUCTO", "PROVEEDOR"])
                                            datos_d = ws_dicc.get_all_values()
                                            fila_a_actualizar = next((i + 1 for i, r in enumerate(datos_d) if len(r)>0 and str(r[0]).strip().upper() == prod_limpio), -1)
                                            if fila_a_actualizar != -1: ws_dicc.update_cell(fila_a_actualizar, 2, prov_limpio)
                                            else: ws_dicc.append_row([prod_limpio, prov_limpio])
                                            st.success(f"✅ ¡Producto {prod_limpio} registrado!")
                                            st.cache_data.clear(); st.rerun()
                                        except Exception as e: st.error(f"🚨 Fallo al guardar: {e}")
                    else:
                        modificar_prov = c_tog2.toggle("✏️ Corregir / Modificar Proveedor")
                        lista_prods_limpia = set([p for p in lista_autorizada if len(p) > 3 and "🛑" not in p])
                        n_prod = c_prod.selectbox("🧪 Producto (Integrado SAP)", sorted(list(lista_prods_limpia)))
                        
                        # 💥 RASTREO DINÁMICO DE MATERIAL (PLANTILLA)
                        mat_item_ing = buscar_codigo_material(n_prod, mapeo_materiales)
                        st.warning(f"🚨 CEBO 2 (Búsqueda Ingresos): Has seleccionado '{n_prod}'. El rastreador encontró: '{mat_item_ing}'")
                        c_mat.text_input("🔢 Cód. Material", value=mat_item_ing, disabled=True)

                        proveedor_asignado = dict_operativo.get(n_prod, "")
                        debe_desbloquear = modificar_prov or not bool(proveedor_asignado.strip())
                        n_prov = c_prov.text_input("🏭 Proveedor", value=proveedor_asignado, disabled=not debe_desbloquear, placeholder="Digite proveedor")
                        
                        if debe_desbloquear:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("💾 GUARDAR PROVEEDOR PERMANENTEMENTE", type="secondary", use_container_width=True):
                                prod_limpio = str(n_prod).strip().upper(); prov_limpio = str(n_prov).strip().upper()
                                if not prov_limpio: st.warning("⚠️ Escribe un proveedor.")
                                elif prov_limpio == proveedor_asignado.upper(): st.info("ℹ️ Proveedor es igual.")
                                else:
                                    with st.spinner("Actualizando Diccionario..."):
                                        try:
                                            gc_temp = inicializar_cliente_gspread()
                                            sh_temp = gc_temp.open_by_url(URL_SHEET_INGRESOS)
                                            try: ws_dicc = sh_temp.worksheet("DICCIONARIO")
                                            except:
                                                ws_dicc = sh_temp.add_worksheet(title="DICCIONARIO", rows="100", cols="2")
                                                ws_dicc.append_row(["PRODUCTO", "PROVEEDOR"])
                                            datos_d = ws_dicc.get_all_values()
                                            fila_a_actualizar = next((i + 1 for i, r in enumerate(datos_d) if len(r)>0 and str(r[0]).strip().upper() == prod_limpio), -1)
                                            if fila_a_actualizar != -1: ws_dicc.update_cell(fila_a_actualizar, 2, prov_limpio)
                                            else: ws_dicc.append_row([prod_limpio, prov_limpio])
                                            st.success("✅ ¡Diccionario Actualizado!")
                                            st.cache_data.clear(); st.rerun()
                                        except Exception as e: st.error(f"🚨 Fallo: {e}")

                with st.expander("⚙️ 2. DATOS OPERATIVOS Y TRAZABILIDAD", expanded=True):
                    col_espacio, col_limpiar = st.columns([3, 1])
                    col_limpiar.button("🧹 VACIAR CASILLAS", on_click=limpiar_campos_operativos, use_container_width=True)
                    f1, f2, f3 = st.columns(3)
                    n_fecha_ing = f2.date_input("🗓️ Fecha de Ingreso a SAP", value=hoy_colombia)
                    semana_calculada = n_fecha_ing.isocalendar()[1]
                    n_semana = f1.text_input("📅 Semana del Año (Auto)", value=str(semana_calculada), disabled=True)
                    n_pista = f3.selectbox("📍 Almacén SAP (Pista)", ["LUCI", "PLUC", "PDIV", "PORI", "TEHO"])
                    f4, f5, f6 = st.columns(3)
                    fk = st.session_state['form_key_m19']
                    n_cant = f4.number_input("⚖️ Cantidad", min_value=0.0, step=1.0, key=f"in_cant_{fk}")
                    n_lote = f5.text_input("📦 Lote", key=f"in_lote_{fk}")
                    n_ff = f6.date_input("⚙️ F. Fabricación (F/F)", value=hoy_colombia)
                    f7, f8, f9, f10 = st.columns(4)
                    n_fv = f7.date_input("⏳ F. Vencimiento (F/V)", value=hoy_colombia)
                    n_factura = f8.text_input("🧾 Factura", key=f"in_factura_{fk}")
                    n_pedido = f9.text_input("🛒 Pedido", key=f"in_pedido_{fk}")
                    n_consecutivo = f10.text_input("🔢 Consecutivo SAP", key=f"in_consecutivo_{fk}")
                    
                    st.markdown("<hr style='margin: 15px 0px; border: 1px solid #d4af37;'>", unsafe_allow_html=True)
                    st.markdown("<p style='color: #0d1b2a; font-size: 14px; font-weight: 900; text-transform: uppercase;'>📋 Panel de Copiado Rápido (1-Clic para SAP)</p>", unsafe_allow_html=True)
                    cp_mat, cp1, cp2, cp3, cp4 = st.columns(5)
                    with cp_mat: st.caption("🔢 MATERIAL"); st.code(mat_item_ing if not es_nuevo_producto else "S/N", language="text")
                    with cp1: st.caption("⚖️ CANTIDAD"); st.code(formatear_numero_sap(n_cant), language="text")
                    with cp2: st.caption("📦 LOTE"); st.code(n_lote if n_lote else "...", language="text")
                    with cp3: st.caption("🧾 FACTURA"); st.code(n_factura if n_factura else "...", language="text")
                    with cp4: st.caption("🛒 PEDIDO"); st.code(n_pedido if n_pedido else "...", language="text")

                    st.markdown("<br>", unsafe_allow_html=True)
                    btn_guardar_nuevo = st.button("🚀 INYECTAR NUEVO LOTE A LA BÓVEDA", type="primary", use_container_width=True)
                    
                    if btn_guardar_nuevo:
                        if not n_prod or str(n_prod).strip() == "": st.error("🚨 El nombre del producto no puede estar vacío.")
                        else:
                            prod_limpio = str(n_prod).strip().upper(); prov_limpio = str(n_prov).strip().upper()
                            if es_nuevo_producto or ((modificar_prov or not bool(proveedor_asignado.strip())) and prov_limpio and prov_limpio != proveedor_asignado.upper()):
                                try:
                                    gc_temp = inicializar_cliente_gspread()
                                    sh_temp = gc_temp.open_by_url(URL_SHEET_INGRESOS)
                                    try: ws_dicc = sh_temp.worksheet("DICCIONARIO")
                                    except:
                                        ws_dicc = sh_temp.add_worksheet(title="DICCIONARIO", rows="100", cols="2")
                                        ws_dicc.append_row(["PRODUCTO", "PROVEEDOR"])
                                    datos_d = ws_dicc.get_all_values()
                                    fila_a_actualizar = next((i + 1 for i, r in enumerate(datos_d) if len(r)>0 and str(r[0]).strip().upper() == prod_limpio), -1)
                                    if fila_a_actualizar != -1: ws_dicc.update_cell(fila_a_actualizar, 2, prov_limpio)
                                    else: ws_dicc.append_row([prod_limpio, prov_limpio])
                                except Exception as e: st.warning(f"Se guardó el diccionario: {e}")

                            # 💥 TÁCTICA APÓSTROFE PARA PROTEGER EL CERO DEL LOTE
                            lote_ing_inject = f"'{str(n_lote).strip()}" if str(n_lote).strip() else ""

                            nueva_fila_drive = []
                            for header in encabezados_limpios_ing:
                                h = header.upper()
                                if "SEMANA" in h: nueva_fila_drive.append(str(semana_calculada))
                                elif "PROV" in h: nueva_fila_drive.append(prov_limpio)
                                elif "INGRESO" in h: nueva_fila_drive.append(n_fecha_ing.strftime("%d/%m/%Y"))
                                elif "PROD" in h: nueva_fila_drive.append(prod_limpio)
                                elif "PISTA" in h: nueva_fila_drive.append(str(n_pista))
                                elif "CANT" in h: nueva_fila_drive.append(str(n_cant))
                                elif "LOTE" in h: nueva_fila_drive.append(lote_ing_inject)
                                elif "F/F" in h: nueva_fila_drive.append(n_ff.strftime("%d/%m/%Y"))
                                elif "F/V" in h: nueva_fila_drive.append(n_fv.strftime("%d/%m/%Y"))
                                elif "FACT" in h: nueva_fila_drive.append(str(n_factura))
                                elif "PEDIDO" in h: nueva_fila_drive.append(str(n_pedido))
                                elif "CONSECUT" in h: nueva_fila_drive.append(str(n_consecutivo))
                                elif "ESTADO" in h: nueva_fila_drive.append("✅ VIGENTE")
                                else: nueva_fila_drive.append("") 
                            
                            try:
                                with st.spinner("Enviando datos con láser matemático..."):
                                    gc_temp = inicializar_cliente_gspread()
                                    sh_temp = gc_temp.open_by_url(URL_SHEET_INGRESOS)
                                    ws_write_ing = sh_temp.worksheets()[0]
                                    try: idx_col_prod = encabezados_limpios_ing.index("PRODUCTO") + 1
                                    except: idx_col_prod = 4 
                                    col_prod_data = ws_write_ing.col_values(idx_col_prod)
                                    last_row_ing = len(col_prod_data)
                                    while last_row_ing > 0 and str(col_prod_data[last_row_ing-1]).strip() == "": last_row_ing -= 1
                                    fila_destino = last_row_ing + 1
                                    rango_inyeccion = f"A{fila_destino}:{get_column_letter(len(encabezados_limpios_ing))}{fila_destino}"
                                    
                                    try: ws_write_ing.update(range_name=rango_inyeccion, values=[nueva_fila_drive], value_input_option='USER_ENTERED')
                                    except: ws_write_ing.update(rango_inyeccion, [nueva_fila_drive], value_input_option='USER_ENTERED')
                                    
                                st.success(f"✅ ¡Lote de {prod_limpio} inyectado exactamente en la fila {fila_destino}!")
                                st.session_state['form_key_m19'] += 1
                                st.cache_data.clear(); st.rerun()
                            except Exception as e: st.error(f"Error al inyectar datos: {e}")

                # --- GENERADOR DE REPORTE CORREO ---
                st.markdown("---")
                st.markdown("### 📧 Reporte Rápido para Correo (Copy & Paste)")
                st.info("💡 **Filtro Anti-Infiltración:** Los registros anulados se ocultan por defecto.")
                
                col_fecha_rep, col_vacia = st.columns([1, 3])
                fecha_reporte = col_fecha_rep.date_input("Fecha a reportar:", value=hoy_colombia)
                
                if col_fecha_ingreso:
                    df['FECHA_ING_TEMP'] = df[col_fecha_ingreso].apply(procesar_fecha_estricta)
                    mask = df['FECHA_ING_TEMP'].apply(lambda x: x.date() if pd.notna(x) else None) == fecha_reporte
                    df_correo = df[mask].copy()
                    if COL_ESTADO in df_correo.columns: df_correo = df_correo[~df_correo[COL_ESTADO].str.contains("ANULADO|ELIMINAR", na=False, case=False)]
                    if not df_correo.empty:
                        df_correo = df_correo.sort_values(by='FILA_EXCEL', ascending=False)
                        df_correo.insert(0, "✅ INCLUIR", True)
                        cols_ed = [c for c in df_correo.columns if c not in ["FILA_EXCEL", "FECHA_ING_TEMP", "FECHA_VENC_DT", COL_ESTADO]]
                        st.markdown("👇 **Paso 1: Desmarca los registros que NO quieres enviar en el correo:**")
                        df_editado_correo = st.data_editor(df_correo[cols_ed], column_config={"✅ INCLUIR": st.column_config.CheckboxColumn("✅ INCLUIR", default=True)}, disabled=[c for c in cols_ed if c != "✅ INCLUIR"], hide_index=True, use_container_width=True, key=f"editor_correo_{fecha_reporte}")
                        df_correo_final = df_editado_correo[df_editado_correo["✅ INCLUIR"] == True].copy()
                        if not df_correo_final.empty:
                            mapa_columnas = {}
                            for col_excel in df_correo_final.columns:
                                c_up = str(col_excel).upper()
                                if "SEMANA" in c_up: mapa_columnas[col_excel] = "SEMANA"
                                elif "PROV" in c_up: mapa_columnas[col_excel] = "PROVEEDOR"
                                elif "INGRESO" in c_up: mapa_columnas[col_excel] = "F. INGRESO"
                                elif "PROD" in c_up: mapa_columnas[col_excel] = "PRODUCTO"
                                elif "PISTA" in c_up: mapa_columnas[col_excel] = "PISTA"
                                elif "CANT" in c_up: mapa_columnas[col_excel] = "CANTIDAD"
                                elif "LOTE" in c_up: mapa_columnas[col_excel] = "LOTE"
                                elif "F/F" in c_up: mapa_columnas[col_excel] = "F/F"
                                elif "F/V" in c_up: mapa_columnas[col_excel] = "F/V"
                                elif "FACT" in c_up: mapa_columnas[col_excel] = "FACTURA"
                                elif "PEDIDO" in c_up: mapa_columnas[col_excel] = "PEDIDO"
                                elif "CONSECUT" in c_up: mapa_columnas[col_excel] = "CONSECUTIVO"
                            df_correo_limpio = df_correo_final[list(mapa_columnas.keys())].rename(columns=mapa_columnas)
                            orden_ideal = ["PROVEEDOR", "PRODUCTO", "PISTA", "CANTIDAD", "LOTE", "F/F", "F/V", "FACTURA", "PEDIDO", "CONSECUTIVO"]
                            df_correo_limpio = df_correo_limpio[[col for col in orden_ideal if col in df_correo_limpio.columns]]
                            html_manual = "<table style='border-collapse: collapse; width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 11px; border: 2px solid #0d1b2a; margin-top: 10px; background-color: #ffffff;'><thead><tr>"
                            for col_name in df_correo_limpio.columns: html_manual += f"<th style='background-color: #0d1b2a; color: #d4af37; padding: 8px 6px; border: 2px solid #0d1b2a; text-align: center; font-weight: 900; text-transform: uppercase; white-space: nowrap;'>{col_name}</th>"
                            html_manual += "</tr></thead><tbody>"
                            for _, row in df_correo_limpio.iterrows():
                                html_manual += "<tr>"
                                for col_name in df_correo_limpio.columns:
                                    val = row[col_name]
                                    val_str = "" if pd.isna(val) or str(val).strip() == "" else (formatear_numero_sap(str(val).strip()) if col_name == "CANTIDAD" else str(val).strip())
                                    if val_str.startswith("'"): val_str = val_str[1:]
                                    html_manual += f"<td style='padding: 8px 6px; border: 1px solid #0d1b2a; text-align: center; color: #000000; font-weight: bold; white-space: nowrap;'>{val_str}</td>"
                                html_manual += "</tr>"
                            html_manual += "</tbody></table>"
                            st.markdown("👇 **Paso 2: Copia la tabla a continuación y pégala en tu correo:**")
                            st.markdown(html_manual, unsafe_allow_html=True)
                        else: st.info("Has desmarcado todos los registros. La tabla final está vacía.")
                    else: st.warning(f"No se encontraron ingresos válidos en la bóveda con la fecha {fecha_reporte.strftime('%d/%m/%Y')}.")

                # --- ESCÁNER DE AUDITORÍA ---
                st.markdown("---")
                st.markdown("<div id='seccion-auditoria'></div>", unsafe_allow_html=True)
                st.markdown("### 🔍 Escáner de Auditoría (Filtros)")
                f_col1, f_col2 = st.columns([1.5, 1])
                filtro_seleccionado = f_col1.radio("Estado Operativo:", ["🌐 Mostrar Todos", "✅ Solo Vigentes", "🚨 Solo Vencidos", "⚠️ Por Vencer (90 Días)"], horizontal=True)
                
                lista_productos_tabla = ["TODOS"] + sorted(list(set([str(x).strip().upper() for x in df[col_producto].dropna() if str(x).strip() != ""]))) if col_producto else ["TODOS"]
                producto_filtro = f_col2.selectbox("🧪 Filtrar por Producto:", lista_productos_tabla)

                st.markdown("<br>", unsafe_allow_html=True)
                f_col3, f_col4, _ = st.columns([1, 1, 1.5])
                fecha_ini_filtro = f_col3.date_input("📅 Ingreso Desde:", value=date(2021, 1, 1))
                fecha_fin_filtro = f_col4.date_input("📅 Ingreso Hasta:", value=hoy_colombia)
                
                df[COL_ESTADO] = df[COL_ESTADO].replace(r'^\s*$', '✅ VIGENTE', regex=True).fillna('✅ VIGENTE')
                df_filtrado = df.copy()
                
                if col_fecha_ingreso:
                    fechas_dt = df_filtrado[col_fecha_ingreso].apply(procesar_fecha_estricta)
                    mask_fechas = fechas_dt.apply(lambda x: x.date() if pd.notna(x) else None)
                    df_filtrado = df_filtrado[(mask_fechas >= fecha_ini_filtro) & (mask_fechas <= fecha_fin_filtro)]

                if producto_filtro != "TODOS" and col_producto: df_filtrado = df_filtrado[df_filtrado[col_producto].str.upper() == producto_filtro]
                if filtro_seleccionado == "✅ Solo Vigentes": df_filtrado = df_filtrado[df_filtrado[COL_ESTADO].str.contains("VIGENTE", case=False, na=False)]
                elif filtro_seleccionado == "🚨 Solo Vencidos" and col_fv: df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO|ELIMINAR", na=False)) & (df_filtrado['FECHA_VENC_DT'] < hoy_ts)]
                elif filtro_seleccionado == "⚠️ Por Vencer (90 Días)" and col_fv: df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO|ELIMINAR", na=False)) & (df_filtrado['FECHA_VENC_DT'] >= hoy_ts) & (df_filtrado['FECHA_VENC_DT'] <= limite_90_dias)]

                if col_fecha_ingreso:
                    df_filtrado['FECHA_SORT'] = df_filtrado[col_fecha_ingreso].apply(procesar_fecha_estricta)
                    df_filtrado = df_filtrado.sort_values(by=['FECHA_SORT', 'FILA_EXCEL'], ascending=[False, False])
                else: df_filtrado = df_filtrado.sort_values(by=['FILA_EXCEL'], ascending=[False])

                st.markdown("### 🛠️ Matriz de Anulaciones (Solo Lectura y Edición de Estado)")
                st.caption("🔒 Haz doble clic en ESTADO/OBSERVACIÓN para anular o ELIMINAR el registro físicamente de la base de datos.")
                
                cols_disabled = [col for col in df_filtrado.columns if col not in [COL_ESTADO, 'FILA_EXCEL', 'FECHA_VENC_DT', 'FECHA_ING_TEMP', 'FECHA_SORT']]
                opciones_estado = ["✅ VIGENTE", "❌ ANULADO: ERROR EN PRECIOS", "❌ ANULADO: ERROR DE CANTIDAD", "❌ ANULADO: DEVOLUCIÓN A PROVEEDOR", "❌ ANULADO: ERROR EN LOTE/FECHAS", "❌ ANULADO: OTRO MOTIVO", "💥 ELIMINAR REGISTRO (BORRADO FÍSICO)"]

                columnas_vista = [c for c in df_filtrado.columns if c not in ['FILA_EXCEL', 'FECHA_VENC_DT', 'FECHA_ING_TEMP', 'FECHA_SORT']]
                df_vista = df_filtrado[columnas_vista].copy()
                
                if "LOTE" in df_vista.columns: df_vista["LOTE"] = df_vista["LOTE"].astype(str).str.lstrip("'")
                
                col_config = {COL_ESTADO: st.column_config.SelectboxColumn("🛡️ ESTADO / OBSERVACIÓN", help="Doble clic para anular o cambiar estado.", width="large", options=opciones_estado, required=True)}
                for c in df_vista.columns:
                    c_up = c.upper()
                    if "SEMANA" in c_up: col_config[c] = st.column_config.TextColumn("📅 SEMANA", width="small")
                    elif "PROV" in c_up: col_config[c] = st.column_config.TextColumn("🏭 PROVEEDOR", width="medium")
                    elif "INGRESO" in c_up: col_config[c] = st.column_config.TextColumn("🗓️ INGRESO SAP", width="medium")
                    elif "PROD" in c_up: col_config[c] = st.column_config.TextColumn("🧪 PRODUCTO", width="large")
                    elif "PISTA" in c_up: col_config[c] = st.column_config.TextColumn("📍 BASE", width="small")
                    elif "CANT" in c_up: col_config[c] = st.column_config.TextColumn("⚖️ CANTIDAD", width="medium")
                    elif "LOTE" in c_up: col_config[c] = st.column_config.TextColumn("📦 LOTE", width="medium")
                    elif "F/F" in c_up: col_config[c] = st.column_config.TextColumn("⚙️ F/F", width="small")
                    elif "F/V" in c_up: col_config[c] = st.column_config.TextColumn("⏳ F/V", width="small")
                    elif "FACT" in c_up: col_config[c] = st.column_config.TextColumn("🧾 FACTURA", width="medium")
                    elif "PEDIDO" in c_up: col_config[c] = st.column_config.TextColumn("🛒 PEDIDO", width="medium")
                    elif "CONSECUT" in c_up: col_config[c] = st.column_config.TextColumn("🔢 CONSECUTIVO", width="medium")
                    
                df_editado = st.data_editor(df_vista, column_config=col_config, disabled=cols_disabled, hide_index=True, use_container_width=True, key="editor_ingresos")

                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 SINCRONIZAR CAMBIOS Y ELIMINACIONES EN DRIVE", type="primary"):
                    cambios = []
                    for i in range(len(df_filtrado)):
                        estado_original = str(df_filtrado.iloc[i][COL_ESTADO]).strip()
                        estado_nuevo = str(df_editado.iloc[i][COL_ESTADO]).strip()
                        if estado_original != estado_nuevo: cambios.append({'fila': int(df_filtrado.iloc[i]['FILA_EXCEL']), 'nuevo': estado_nuevo})
                    
                    if cambios:
                        eliminaciones = [c for c in cambios if "ELIMINAR REGISTRO" in c['nuevo']]
                        actualizaciones = [c for c in cambios if "ELIMINAR REGISTRO" not in c['nuevo']]
                        cambios_exitosos = False
                        gc_temp = inicializar_cliente_gspread()
                        sh_temp = gc_temp.open_by_url(URL_SHEET_INGRESOS)
                        ws_write_ing = sh_temp.worksheets()[0]
                        
                        for act in actualizaciones:
                            try: ws_write_ing.update_cell(act['fila'], idx_col_estado, act['nuevo']); cambios_exitosos = True
                            except Exception as e: st.error(f"Error al actualizar fila {act['fila']}: {e}")
                                    
                        if eliminaciones:
                            for eli in sorted(eliminaciones, key=lambda x: x['fila'], reverse=True):
                                try: ws_write_ing.delete_row(eli['fila']); cambios_exitosos = True
                                except AttributeError:
                                    try: ws_write_ing.delete_rows(eli['fila']); cambios_exitosos = True
                                    except Exception as e: st.error(f"Error fatal API Google. Fila {eli['fila']}. {e}")
                                except Exception as e: st.error(f"Error al eliminar fila {eli['fila']}. {e}")
                        if cambios_exitosos:
                            st.success("✅ ¡Misión Cumplida! Base de datos sincronizada y purgada exitosamente.")
                            st.cache_data.clear(); st.rerun()
                    else: st.info("No se detectaron cambios ni órdenes de eliminación.")

                st.markdown("---")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_editado.to_excel(writer, sheet_name='Auditoria_Ingresos', index=False)
                    ws_excel = writer.sheets['Auditoria_Ingresos']
                    header_font, header_fill = Font(bold=True, color="FFFFFF"), PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                    for cell in ws_excel[1]: cell.font = header_font; cell.fill = header_fill; cell.alignment = Alignment(horizontal='center', vertical='center')
                    for col in ws_excel.columns:
                        max_length = 0
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length: max_length = len(cell.value)
                            except: pass
                        ws_excel.column_dimensions[col[0].column_letter].width = (max_length + 2)

                st.download_button("💾 DESCARGAR REPORTE DE AUDITORÍA (EXCEL)", data=buffer.getvalue(), file_name=f"Reporte_Auditoria_Ingresos_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ========================================================================
    # 🚚 PESTAÑA 2: MOVIMIENTOS INTERNOS (TRASLADOS)
    # ========================================================================
    with tab_traslados:
        st.markdown(f"<a href='{URL_SHEET_TRASLADOS}' target='_blank' class='btn-ascensor' style='background-color:#2F75B5; border-color:#1d4e7a; color:#ffffff !important;'>👁️ VER BASE DE TRASLADOS EN GOOGLE SHEETS</a>", unsafe_allow_html=True)
        st.write("Panel táctico de logística interna. Registra los movimientos de inventario entre las diferentes bases operativas.")

        if not df_traslados.empty: st.markdown(f"<div class='kpi-card kpi-azul'><div class='kpi-titulo'>🚚 Movimientos Históricos Registrados</div><p class='kpi-valor'>{len(df_traslados)}</p></div>", unsafe_allow_html=True)
        else: st.markdown(f"<div class='kpi-card kpi-azul'><div class='kpi-titulo'>🚚 Movimientos Históricos Registrados</div><p class='kpi-valor'>0</p></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### ➕ Inyector de Movimiento Interno")

        with st.expander("📍 1. DATOS DEL TRASLADO", expanded=True):
            col_espacio_t, col_limpiar_t = st.columns([3, 1])
            col_limpiar_t.button("🧹 VACIAR CASILLAS", on_click=limpiar_campos_traslados, key="btn_limpiar_traslados", use_container_width=True)

            t1, t2, t3 = st.columns(3)
            fk_t = st.session_state['form_key_m19_traslados']
            t_fecha = t2.date_input("🗓️ Fecha de Traslado", value=hoy_colombia, key=f"t_fecha_{fk_t}")
            semana_traslado = t_fecha.isocalendar()[1]
            t_semana = t1.text_input("📅 Semana del Año", value=str(semana_traslado), disabled=True, key=f"t_semana_{fk_t}")
            t_consecutivo = t3.text_input("🔢 Consecutivo", key=f"t_consecutivo_{fk_t}")

            st.markdown("<hr style='margin: 10px 0px; border: 1px solid #e2e8f0;'>", unsafe_allow_html=True)
            
            pistas_disponibles = ["LUCI", "PLUC", "PDIV", "PORI", "TEHO"]
            p1, p2 = st.columns(2)
            t_origen = p1.selectbox("🛫 Pista Origen", pistas_disponibles, index=0, key=f"t_origen_{fk_t}")
            t_destino = p2.selectbox("🛬 Pista Destino", pistas_disponibles, index=1 if len(pistas_disponibles) > 1 else 0, key=f"t_destino_{fk_t}")

            st.markdown("<hr style='margin: 10px 0px; border: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

            lista_prods_limpia_t = set([p for p in lista_autorizada if len(p) > 3 and "🛑" not in p])
            lista_prods_ordenada_t = sorted(list(lista_prods_limpia_t))

            tr1, tr_mat, tr2, tr3 = st.columns([2, 1, 1, 1])
            t_producto = tr1.selectbox("🧪 Producto a Trasladar", lista_prods_ordenada_t, key=f"t_producto_{fk_t}")
            
            # 💥 RASTREO DINÁMICO DE MATERIAL
            mat_item_tras = buscar_codigo_material(t_producto, mapeo_materiales)
            st.warning(f"🚨 CEBO 3 (Búsqueda Traslados): Has seleccionado '{t_producto}'. El rastreador encontró: '{mat_item_tras}'")
            tr_mat.text_input("🔢 Cód. Material", value=mat_item_tras, disabled=True, key=f"t_mat_{fk_t}")

            t_cantidad = tr2.number_input("⚖️ Cantidad", min_value=0.0, step=1.0, key=f"t_cantidad_{fk_t}")
            t_unidad = tr3.selectbox("📦 Unidad", ["LITROS", "KILOS", "GALONES", "UNIDADES"], key=f"t_unidad_{fk_t}")

            tr4, tr5 = st.columns(2)
            opciones_obs = ["SIN NOVEDAD", "ANULACIÓN", "TRANSFORMACIÓN DE LOTE", "OTRO"]
            t_observacion_sel = tr4.selectbox("📝 Observación", opciones_obs, key=f"t_obs_sel_{fk_t}")
            t_lote = tr5.text_input("📦 Lote", key=f"t_lote_{fk_t}")

            if t_observacion_sel == "OTRO": t_observacion = st.text_input("📝 Especifique la observación:", key=f"t_obs_otro_{fk_t}")
            else: t_observacion = t_observacion_sel

            st.markdown("<hr style='margin: 15px 0px; border: 1px solid #d4af37;'>", unsafe_allow_html=True)
            st.markdown("<p style='color: #0d1b2a; font-size: 14px; font-weight: 900; text-transform: uppercase;'>📋 Panel de Copiado Rápido (1-Clic para SAP)</p>", unsafe_allow_html=True)

            cpt_mat, cpt1, cpt2, cpt3, cpt4 = st.columns(5)
            with cpt_mat: st.caption("🔢 MATERIAL"); st.code(mat_item_tras, language="text")
            with cpt1: st.caption("⚖️ CANTIDAD"); st.code(formatear_numero_sap(t_cantidad), language="text")
            with cpt2: st.caption("📦 LOTE"); st.code(t_lote if t_lote else "...", language="text")
            with cpt3: st.caption("🛫 ORIGEN"); st.code(t_origen, language="text")
            with cpt4: st.caption("🛬 DESTINO"); st.code(t_destino, language="text")

            st.markdown("<br>", unsafe_allow_html=True)
            btn_guardar_traslado = st.button("🚀 REGISTRAR TRASLADO EN LA BÓVEDA", type="primary", use_container_width=True)

            if btn_guardar_traslado:
                if not t_consecutivo.strip(): st.error("🚨 Debes ingresar un número de Consecutivo.")
                elif t_cantidad <= 0: st.error("🚨 La cantidad debe ser mayor a cero.")
                elif t_origen == t_destino: st.error("🚨 La pista de origen y destino no pueden ser la misma.")
                else:
                    try:
                        with st.spinner("Enviando traslado a la nube..."):
                            pista_combinada = f"{t_origen}-{t_destino}"
                            fecha_str = t_fecha.strftime("%d/%m/%Y")
                            cantidad_formateada = formatear_numero_sap(t_cantidad)

                            # 💥 TÁCTICA APÓSTROFE PARA PROTEGER EL CERO DEL LOTE
                            lote_tras_inject = f"'{str(t_lote).strip()}" if str(t_lote).strip() else ""
                            
                            nueva_fila_traslado = []
                            for h in encabezados_limpios_tras:
                                h_up = h.upper()
                                if "CONSECUTIVO" in h_up: nueva_fila_traslado.append(str(t_consecutivo).strip())
                                elif "FECHA" in h_up: nueva_fila_traslado.append(fecha_str)
                                elif "PROD" in h_up: nueva_fila_traslado.append(str(t_producto).upper().strip())
                                elif "CANT" in h_up: nueva_fila_traslado.append(cantidad_formateada)
                                elif "UNIDAD" in h_up: nueva_fila_traslado.append(str(t_unidad).upper())
                                elif "PISTA" in h_up: nueva_fila_traslado.append(pista_combinada)
                                elif "SEMANA" in h_up: nueva_fila_traslado.append(str(semana_traslado))
                                elif "OBSERVAC" in h_up: nueva_fila_traslado.append(str(t_observacion).strip())
                                elif "LOTE" in h_up: nueva_fila_traslado.append(lote_tras_inject)
                                else: nueva_fila_traslado.append("")
                            
                            gc_temp = inicializar_cliente_gspread()
                            sh_temp = gc_temp.open_by_url(URL_SHEET_TRASLADOS)
                            ws_write = sh_temp.worksheets()[0]

                            col_a = ws_write.col_values(1)
                            last_row = len(col_a)
                            while last_row > 0 and str(col_a[last_row-1]).strip() == "": last_row -= 1
                                
                            fila_destino = last_row + 1
                            rango_inyeccion = f"A{fila_destino}:{get_column_letter(len(nueva_fila_traslado))}{fila_destino}"
                            
                            try: ws_write.update(range_name=rango_inyeccion, values=[nueva_fila_traslado], value_input_option='USER_ENTERED')
                            except: ws_write.update(rango_inyeccion, [nueva_fila_traslado], value_input_option='USER_ENTERED')
                            
                        st.success(f"✅ ¡Traslado de {t_producto} registrado con éxito en la fila {fila_destino}!")
                        st.session_state['form_key_m19_traslados'] += 1
                        st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"Error al registrar traslado: {e}")

        # --- VISOR HISTÓRICO DE TRASLADOS ---
        st.markdown("---")
        st.markdown("### 📋 Visor y Edición Histórica de Movimientos")
        
        if not df_traslados.empty:
            df_traslados_vista = df_traslados.copy()
            
            st.markdown("#### 🔍 Escáner de Filtrado")
            col_prod_t = next((c for c in df_traslados_vista.columns if "PRODUCTO" in c), None)
            
            if col_prod_t:
                productos_puros_t = set([str(x).strip().upper() for x in df_traslados_vista[col_prod_t].dropna() if str(x).strip() != ""])
                lista_productos_tabla_t = ["TODOS"] + sorted(list(productos_puros_t))
            else: lista_productos_tabla_t = ["TODOS"]
            
            f_col_t1, f_col_t2, f_col_t3 = st.columns([1.5, 2, 1])
            producto_filtro_t = f_col_t1.selectbox("🧪 Filtrar por Producto:", lista_productos_tabla_t, key="filtro_t_prod")
            st.markdown("<br>", unsafe_allow_html=True)
            
            if producto_filtro_t != "TODOS" and col_prod_t: df_traslados_vista = df_traslados_vista[df_traslados_vista[col_prod_t].str.upper() == producto_filtro_t]

            df_traslados_vista.insert(0, "🛡️ ACCIÓN", "✅ MANTENER")

            if "FECHA" in df_traslados_vista.columns:
                df_traslados_vista['FECHA_SORT'] = df_traslados_vista['FECHA'].apply(procesar_fecha_estricta)
                df_traslados_vista = df_traslados_vista.sort_values(by=['FECHA_SORT', 'FILA_EXCEL'], ascending=[False, False])
            else: df_traslados_vista = df_traslados_vista.sort_values(by=['FILA_EXCEL'], ascending=[False])

            st.caption("🔒 Haz doble clic en la columna '🛡️ ACCIÓN' para marcar y **ELIMINAR** un traslado de la Bóveda de Google Sheets.")

            columnas_vista_t = [c for c in df_traslados_vista.columns if c not in ['FILA_EXCEL', 'FECHA_SORT']]
            df_vista_t = df_traslados_vista[columnas_vista_t].copy()
            
            if "LOTE" in df_vista_t.columns: df_vista_t["LOTE"] = df_vista_t["LOTE"].astype(str).str.lstrip("'")
            
            cols_disabled_t = [c for c in df_vista_t.columns if c != "🛡️ ACCIÓN"]
            col_config_t = {"🛡️ ACCIÓN": st.column_config.SelectboxColumn("🛡️ ACCIÓN", help="Selecciona ELIMINAR para borrar esta fila.", options=["✅ MANTENER", "💥 ELIMINAR REGISTRO"], required=True)}
            
            for c in df_vista_t.columns:
                if c == "🛡️ ACCIÓN": continue
                c_up = c.upper()
                if "SEMANA" in c_up: col_config_t[c] = st.column_config.TextColumn("📅 SEM", width="small")
                elif "FECHA" in c_up: col_config_t[c] = st.column_config.TextColumn("🗓️ FECHA", width="medium")
                elif "PROD" in c_up: col_config_t[c] = st.column_config.TextColumn("🧪 PRODUCTO", width="large")
                elif "PISTA" in c_up: col_config_t[c] = st.column_config.TextColumn("📍 RUTA", width="medium")
                elif "CANT" in c_up: col_config_t[c] = st.column_config.TextColumn("⚖️ CANTIDAD", width="medium")
                elif "LOTE" in c_up: col_config_t[c] = st.column_config.TextColumn("📦 LOTE", width="medium")
                elif "OBSER" in c_up: col_config_t[c] = st.column_config.TextColumn("📝 OBS", width="medium")
                elif "CONSECUT" in c_up: col_config_t[c] = st.column_config.TextColumn("🔢 CONSECUTIVO", width="medium")

            df_editado_t = st.data_editor(df_vista_t, column_config=col_config_t, disabled=cols_disabled_t, hide_index=True, use_container_width=True, key="editor_traslados")

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 EJECUTAR ELIMINACIÓN DE TRASLADOS EN DRIVE", type="primary", key="btn_del_traslados"):
                eliminaciones = [int(df_traslados_vista.iloc[i]['FILA_EXCEL']) for i in range(len(df_traslados_vista)) if "ELIMINAR REGISTRO" in str(df_editado_t.iloc[i]["🛡️ ACCIÓN"]).strip()]

                if eliminaciones:
                    cambios_exitosos = False
                    gc_temp = inicializar_cliente_gspread()
                    sh_temp = gc_temp.open_by_url(URL_SHEET_TRASLADOS)
                    ws_t = sh_temp.worksheet(titulo_ws_traslados)

                    for eli in sorted(eliminaciones, reverse=True):
                        try: ws_t.delete_row(eli); cambios_exitosos = True
                        except AttributeError:
                            try: ws_t.delete_rows(eli); cambios_exitosos = True
                            except Exception as e: st.error(f"Error API Fila {eli}: {e}")
                        except Exception as e: st.error(f"Error al eliminar fila {eli}: {e}")

                    if cambios_exitosos:
                        st.success("✅ ¡Objetivo neutralizado! El traslado ha sido borrado del sistema.")
                        st.cache_data.clear(); st.rerun()
                else: st.info("ℹ️ No marcaste ninguna fila con la acción de '💥 ELIMINAR REGISTRO'.")

    st.markdown("""<a href="#inicio-modulo-19" class="btn-ascensor" style="margin-top: 20px;">👆 VOLVER AL INICIO (ARRIBA) 👆</a>""", unsafe_allow_html=True)

if __name__ == "__main__":
    ejecutar()
