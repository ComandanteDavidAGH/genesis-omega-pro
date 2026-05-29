import pandas as pd
import streamlit as st
import io
import json
import re
import unicodedata
from datetime import datetime
import dateutil.parser
import modulos.m0_centro_mando as m0
import modulos.m8_reporte_hectareas as m8
import modulos.m9_dashboard_tactico as m9
import modulos.m7_arqueo_inventarios as m7

# Aquí definimos quiénes tienen acceso al sistema extrayendo claves de la bóveda
USUARIOS_CREDENTIALS = {
    "usernames": {
        "comandante": {
            "name": "Comandante Omega",
            "password": st.secrets["passwords"]["comandante"], 
            "role": "ADMIN"
        },
        "gerencia": {
            "name": "Visor Gerencial / Cliente",
            "password": st.secrets["passwords"]["gerencia"], 
            "role": "VIEWER"
        }
    }
}

# Inicializar la memoria de sesión para el inicio de sesión
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False
if 'usuario_rol' not in st.session_state:
    st.session_state['usuario_rol'] = None
if 'usuario_nombre' not in st.session_state:
    st.session_state['usuario_nombre'] = None
    
# Imports de conexiones y apis
import openpyxl
import gspread
import plotly.express as px
import plotly.graph_objects as go  # <-- AGREGUE ESTA LÍNEA AQUÍ

# Intentar importar matplotlib para el mapa de calor, si falla, el sistema sigue
try:
    import matplotlib
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 🔐 CONTROL DE ACCESO PERIMETRAL ---
if not st.session_state['autenticado']:
    st.markdown("<h2 style='text-align: center; color: #0d1b2a;'>🚀 GÉNESIS OMEGA PRO</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Ingrese sus coordenadas de acceso para activar los radares.</p>", unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 2, 1])
    with col_log2:
        with st.form("Formulario de Autenticación"):
            user_input = st.text_input("🛰️ Usuario:")
            pass_input = st.text_input("🔑 Contraseña:", type="password")
            btn_login = st.form_submit_button("🔓 ACTIVAR SISTEMA", use_container_width=True)
            
            if btn_login:
                if user_input in USUARIOS_CREDENTIALS["usernames"]:
                    datos_user = USUARIOS_CREDENTIALS["usernames"][user_input]
                    if pass_input == datos_user["password"]:
                        st.session_state['autenticado'] = True
                        st.session_state['usuario_rol'] = datos_user["role"]
                        st.session_state['usuario_nombre'] = datos_user["name"]
                        st.success(f"🔓 Acceso Concedido. Bienvenido {datos_user['name']}")
                        st.rerun()
                    else:
                        st.error("🚨 Contraseña incorrecta. Intento registrado.")
                else:
                    st.error("🚨 Usuario no identificado en el perímetro.")
    st.stop() # DETIENE EL CÓDIGO AQUÍ SI NO ESTÁ AUTENTICADO

# =====================================================================
# si el código llega aquí, significa que el usuario ya se autenticó
# =====================================================================

# --- 2. ARTILLERÍA VISUAL Y CSS ---
arsenal_css = """
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}

/* Botones Principales */
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }

/* 🎯 CORRECCIÓN: Botones Secundarios Generales (Pantalla Clara) */
button[kind="secondary"] { background-color: transparent !important; color: #0d1b2a !important; border: 1px solid #0d1b2a !important; transition: 0.3s; }
button[kind="secondary"]:hover { background-color: #0d1b2a !important; color: #d4af37 !important; }

/* 🎯 EXCEPCIÓN TÁCTICA: Botones en la Barra Lateral Oscura (Cerrar Sesión) */
[data-testid="stSidebar"] button[kind="secondary"] { color: white !important; border: 1px solid #d4af37 !important; }
[data-testid="stSidebar"] button[kind="secondary"]:hover { background-color: #d4af37 !important; color: #0d1b2a !important; }

div[data-baseweb="input"] input, div[data-baseweb="select"] { color: black !important; background-color: white !important; font-weight: bold; }
th { background-color: #f0f2f6 !important; color: black !important; }
</style>
"""
st.markdown(arsenal_css, unsafe_allow_html=True)

# --- 3. FUNCIONES GLOBALES TÁCTICAS ---
def purificar_lote(lote):
    if pd.isna(lote) or lote is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(lote).upper().strip())

def quitar_tildes(s):
    if pd.isna(s) or s is None: return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(s).upper().strip()) if unicodedata.category(c) != 'Mn')

def extraer_numero(valor):
    if pd.isna(valor) or valor == "": return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    v = str(valor).strip().upper().replace("$", "").replace(" ", "")
    v = re.sub(r'[^\d.,-]', '', v)
    if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    try: return float(v)
    except: return 0.0

def fmt_sap(val): 
    return f"{int(round(val, 0)):,}".replace(",", ".")

def limpiar_texto_vba(t):
    if t is None: return ""
    temp = str(t).upper().strip()
    temp = temp.replace(chr(160), " ").replace(".", "")
    while "  " in temp: temp = temp.replace(" ", " ")
    return temp

def val_seguro(v):
    try: return float(v)
    except: return 0.0

def limpiar_val_dom(v):
    if v is None: return 0.0
    s = str(v).strip()
    if s in ["", "-"]: return 0.0 
    try:
        s = s.replace('$', '').replace(' ', '').replace(',', '.')
        return float(s)
    except: return 0.0

def procesar_fecha_pesada(v):
    if not v or str(v).strip() == "": return None
    try:
        if isinstance(v, (int, float)):
            f = datetime(1899, 12, 30) + pd.Timedelta(days=int(v))
            return f if f.year > 2020 else None
        v_str = str(v).lower().strip()
        if v_str.replace('.', '').isdigit():
            f = datetime(1899, 12, 30) + pd.Timedelta(days=int(float(v_str)))
            return f if f.year > 2020 else None
        meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
        for mes, num_mes in meses.items():
            if mes in v_str:
                match_ano = re.search(r'\d{4}', v_str)
                match_dia = re.search(r'\b\d{1,2}\b', v_str)
                if match_ano and match_dia:
                    f = datetime(int(match_ano.group()), num_mes, int(match_dia.group()))
                    return f if f.year > 2020 else None
        if "/" in v_str or "-" in v_str:
            f = dateutil.parser.parse(v_str, dayfirst=True)
            return f if f.year > 2020 else None
    except: pass
    return None

# =====================================================================
# --- 3.5 🛡️ MOTOR DE CACHÉ Y CONEXIÓN SATELITAL (ANTIBLOQUEOS) ---
# =====================================================================

@st.cache_resource(show_spinner=False)
def conectar_satelite():
    """Abre la conexión a Google una sola vez por sesión."""
    if "gcp_credentials" in st.secrets:
        return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
    else:
        return gspread.service_account(filename='credenciales.json')

@st.cache_data(show_spinner=False, ttl=1800)
def descargar_matriz_rapida(url, pestaña):
    """Descarga los miles de datos y los guarda en RAM ultrarrápida."""
    try:
        gc = conectar_satelite()
        boveda = gc.open_by_url(url)
        try:
            hoja = boveda.worksheet(pestaña)
        except:
            if "TABLA 1" in pestaña.upper():
                hoja = next((s for s in boveda.worksheets() if "TABLA 1" in s.title.upper()), boveda.sheet1)
            else:
                hoja = boveda.worksheet(pestaña)
        return hoja.get_all_values(value_render_option='UNFORMATTED_VALUE')
    except Exception as e:
        st.error(f"🚨 Falla en el satélite al descargar {pestaña}: {e}")
        return []

# --- 4. MENÚ MAESTRO (CUARTEL GENERAL) ---
with st.sidebar:
    st.markdown(f"<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
    st.markdown(f"<p style='text-align: center; color: white; font-size:12px;'>👤 {st.session_state['usuario_nombre']}</p>", unsafe_allow_html=True)
    
    # BOTÓN DE CIERRE DE SESIÓN TÁCTICO
    if st.button("🔒 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.session_state['usuario_rol'] = None
        st.session_state['usuario_nombre'] = None
        st.rerun()
        
    st.markdown("---")
    
    # 🕵️‍♂️ REGLA DE RANGO: Si es ADMIN (Usted), ve todo. Si es VIEWER (Gerente/Cliente), solo ve el Dashboard
    if st.session_state['usuario_rol'] == "ADMIN":
        if st.button("🔄 Cargar Cócteles / Aviones", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
            
        menu = st.radio("🛰️ SELECCIONE LA OPERACIÓN:", [
            "🏠 Centro de Mando", 
            "🛠️ 1. Mantenimiento Plantilla SAP",
            "📥 2. Carga Facturación", 
            "⚙️ 3. Validación de Misión", 
            "⌨️ 4. Ingreso Manual Acelerado (OS)", 
            "📈 5. Sincronización Precios",
            "✈️ 6. Rastreo Dominicales",
            "⚖️ 7. Arqueo de Inventarios",
            "📊 8. Reporte Hectáreas (Pistas)",
            "📈 9. Dashboard Táctico",
            "📊 10. Inteligencia de Costos (BI)"
        ])
    else:
        # El gerente o cliente NO tiene opciones operativas, va directo al Dashboard
        menu = "📈 9. Dashboard Táctico"
        st.info("🛰️ Modo Consulta Gerencial Activado. Acceso restringido a reportes visuales.")

    st.info(f"📅 Operación: {datetime.now().strftime('%Y-%m-%d')}")
    

# =====================================================================
# 🏠 0. CENTRO DE MANDO
# =====================================================================
if menu == "🏠 Centro de Mando":
    m0.renderizar()
# =====================================================================
# 🛠️ 1. MANTENIMIENTO PLANTILLA SAP (CON SINCRONIZADOR INTELIGENTE)
# =====================================================================
elif menu == "🛠️ 1. Mantenimiento Plantilla SAP":
    st.markdown("<h1 class='titulo-principal'>Inteligencia de Precios SAP</h1>", unsafe_allow_html=True)
    
    f_sap_raw = st.file_uploader("📥 1. Suba la Sábana Cruda de SAP", type=["xlsx", "xls", "csv"])
    
    if f_sap_raw:
        if st.button("🚀 PASO A: PURIFICAR Y CARGAR A PLANTILLA", type="primary", use_container_width=True):
            with st.spinner("Ejecutando protocolo Samurai..."):
                try:
                    nombre_archivo = f_sap_raw.name.lower()
                    if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xls'):
                        df = pd.read_excel(f_sap_raw)
                    else:
                        try:
                            df = pd.read_csv(f_sap_raw, sep=None, engine='python', encoding='utf-8')
                        except:
                            f_sap_raw.seek(0)
                            df = pd.read_csv(f_sap_raw, sep=None, engine='python', encoding='latin1')
                    
                    df = df.dropna(subset=[df.columns[0]])
                    df = df[~df.iloc[:, 0].astype(str).str.contains('\*')]
                    if len(df.columns) >= 11:
                        df = df.sort_values(by=df.columns[10], ascending=True)
                    
                    df_final = df.iloc[:, 0:9].copy()
                    df_final['J'] = df.iloc[:, 10].values
                    unicos = sorted(df.iloc[:, 10].astype(str).unique().tolist())
                    
                    if "gcp_credentials" in st.secrets:
                        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                        
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    hoja_plantilla = boveda.worksheet("Plantilla")
                    hoja_plantilla.batch_clear(["A3:K5000"])
                    hoja_plantilla.update("A3", df_final.fillna("").values.tolist(), value_input_option='USER_ENTERED')
                    hoja_plantilla.update("K3", [[x] for x in unicos], value_input_option='USER_ENTERED')
                    
                    st.success("✅ PASO A COMPLETADO: Datos frescos cargados en Plantilla.")
                    st.session_state['paso_a_listo'] = True
                except Exception as e:
                    st.error(f"🚨 Error en Paso A: {e}")

        st.markdown("---")
        st.markdown("### ⚡ PASO B: SINCRONIZADOR DE PRECIOS (ESTADO DEL ARSENAL)")
        
        if st.button("🔍 ESCANEAR ESTADO ACTUAL", use_container_width=True):
            try:
                if "gcp_credentials" in st.secrets:
                    gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                else:
                    gc = gspread.service_account(filename='credenciales.json')
                
                url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                sh = gc.open_by_url(url_boveda)
                ws_conf = sh.worksheet("Configuración")
                
                data = ws_conf.get_all_values()
                df_conf = pd.DataFrame(data[1:], columns=data[0])
                
                radar = df_conf.iloc[:, [8, 9, 10]].copy()
                radar.columns = ['PRODUCTO', 'PRECIO_ACTUAL', 'PRECIO_SAP']
                
                radar['PRECIO_ACTUAL'] = radar['PRECIO_ACTUAL'].apply(extraer_numero)
                radar['PRECIO_SAP'] = radar['PRECIO_SAP'].apply(extraer_numero)
                radar['DIFERENCIA'] = (radar['PRECIO_SAP'] - radar['PRECIO_ACTUAL']).round(2)
                radar['ESTADO'] = radar['DIFERENCIA'].apply(lambda x: "✅ OK" if x == 0 else "❌ DESFASE")
                radar = radar.sort_values(by="ESTADO", ascending=False)
                
                st.markdown("#### 🛰️ Reporte de Situación:")
                def color_estado(val):
                    if val == "✅ OK": return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'
                    if val == "❌ DESFASE": return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center;'
                    return ''

                st.dataframe(radar.style.map(color_estado, subset=['ESTADO']), use_container_width=True, hide_index=True)
                
                hay_desfase = (radar['ESTADO'] == "❌ DESFASE").any()
                if not hay_desfase:
                    st.success("🟢 TODO EL SISTEMA ESTÁ EN NIVEL 'OK'. No se requieren ajustes.")
                else:
                    st.warning("⚠️ SE DETECTARON DESFASES. Proceda a la inyección para nivelar.")
                    st.session_state['datos_para_sincronizar'] = True

            except Exception as e:
                st.error(f"Error al escanear: {e}")

        if st.session_state.get('datos_para_sincronizar'):
            st.markdown("---")
            if st.button("✅ APROBAR E INYECTAR PRECIOS (MODO SEGURO)", type="primary", use_container_width=True):
                with st.spinner("Inyectando quirúrgicamente Columna K en Columna J..."):
                    try:
                        if "gcp_credentials" in st.secrets:
                            gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                        else:
                            gc = gspread.service_account(filename='credenciales.json')
                        
                        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                        ws_conf = sh.worksheet("Configuración")
                        data_full = ws_conf.get_all_values()
                        
                        valores_para_j = []
                        for fila in data_full[1:]:
                            valor_k = fila[10] if len(fila) > 10 else ""
                            valores_para_j.append([valor_k])
                        
                        if valores_para_j:
                            rango_destino = f"J2:J{len(valores_para_j) + 1}"
                            ws_conf.update(rango_destino, valores_para_j, value_input_option='USER_ENTERED')
                        
                        st.balloons()
                        st.success(f"🎯 INYECCIÓN EXITOSA. Se actualizaron {len(valores_para_j)} celdas en la columna J.")
                        del st.session_state['datos_para_sincronizar']
                    except Exception as e:
                        st.error(f"🚨 FALLA EN LA INYECCIÓN: {e}")

# =====================================================================
# 📥 2. CARGA FACTURACIÓN
# =====================================================================
elif menu == "📥 2. Carga Facturación":
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Facturación</h1>", unsafe_allow_html=True)
    
    import io
    
    # ---------------------------------------------------------
    # 🧠 SISTEMA DE MEMORIA PERSISTENTE PARA ARCHIVOS SAP
    # ---------------------------------------------------------
    if 'mem_sabana' not in st.session_state: st.session_state['mem_sabana'] = None
    if 'name_sabana' not in st.session_state: st.session_state['name_sabana'] = None
    if 'mem_pedidos' not in st.session_state: st.session_state['mem_pedidos'] = None
    if 'name_pedidos' not in st.session_state: st.session_state['name_pedidos'] = None

    c1, c2, c3 = st.columns(3)
    
    # 📦 CAJA FUERTE 1: SÁBANA
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        if st.session_state['mem_sabana'] is None:
            f_sabana_up = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv"], key="sab")
            if f_sabana_up:
                st.session_state['mem_sabana'] = f_sabana_up.getvalue()
                st.session_state['name_sabana'] = f_sabana_up.name
                st.rerun()
        else:
            st.success(f"✅ Sábana en memoria: {st.session_state['name_sabana']}")
            if st.button("🔄 Cambiar Sábana", use_container_width=True):
                st.session_state['mem_sabana'] = None
                st.rerun()

    # 📦 CAJA FUERTE 2: PEDIDOS
    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        if st.session_state['mem_pedidos'] is None:
            f_pedidos_up = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv"], key="ped")
            if f_pedidos_up:
                st.session_state['mem_pedidos'] = f_pedidos_up.getvalue()
                st.session_state['name_pedidos'] = f_pedidos_up.name
                st.rerun()
        else:
            st.success(f"✅ Pedidos en memoria: {st.session_state['name_pedidos']}")
            if st.button("🔄 Cambiar Pedidos", use_container_width=True):
                st.session_state['mem_pedidos'] = None
                st.rerun()

    # 🚁 CAJA FUERTE 3: PISTAS (Este cambia cada vez, lo dejamos igual)
    with c3:
        st.markdown("### 🚁 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pis")

    # 🪄 TRUCO DE MAGIA: Reconstruimos los archivos para que el botón procesador no note la diferencia
    f_sabana = None
    if st.session_state['mem_sabana']:
        f_sabana = io.BytesIO(st.session_state['mem_sabana'])
        f_sabana.name = st.session_state['name_sabana']
        
    f_pedidos = None
    if st.session_state['mem_pedidos']:
        f_pedidos = io.BytesIO(st.session_state['mem_pedidos'])
        f_pedidos.name = st.session_state['name_pedidos']

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas:
            with st.spinner("Sincronizando los 3 frentes..."):
                try: 
                    nombre_sabana = f_sabana.name.lower()
                    if nombre_sabana.endswith(('.xlsx', '.xls')):
                        st.session_state['df_sabana'] = pd.read_excel(f_sabana)
                    else:
                        try:
                            st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='utf-8')
                        except UnicodeDecodeError:
                            f_sabana.seek(0)
                            st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='latin1')
                    
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos) if f_pedidos.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    if "gcp_credentials" in st.secrets:
                        cred_dict = dict(st.secrets["gcp_credentials"])
                        gc = gspread.service_account_from_dict(cred_dict)
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                    
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    
                    st.session_state['df_config'] = pd.DataFrame(boveda.worksheet("TABLA 2").get_all_values()[1:], columns=boveda.worksheet("TABLA 2").get_all_values()[0])
                    st.session_state['df_mezclas'] = pd.DataFrame(boveda.worksheet("DD_Mesclas").get_all_values()[1:], columns=boveda.worksheet("DD_Mesclas").get_all_values()[0])
                    st.session_state['df_config_base'] = pd.DataFrame(boveda.worksheet("Configuración").get_all_values()[1:], columns=boveda.worksheet("Configuración").get_all_values()[0])
                    
                    hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023") 
                    datos_apoyo = hoja_apoyo.get_all_values()
                    
                    fila_titulos = 0
                    for i, fila in enumerate(datos_apoyo[:20]):
                        if any('FINCA' in str(celda).upper() for celda in fila):
                            fila_titulos = i
                            break
                            
                    encabezados_crudos = datos_apoyo[fila_titulos]
                    encabezados_limpios = []
                    vistos = {}
                    for col in encabezados_crudos:
                        col_str = str(col).strip()
                        if col_str == "": col_str = "Vacio"
                        if col_str in vistos:
                            vistos[col_str] += 1
                            encabezados_limpios.append(f"{col_str}_{vistos[col_str]}")
                        else:
                            vistos[col_str] = 0
                            encabezados_limpios.append(col_str)
                            
                    st.session_state['df_apoyo'] = pd.DataFrame(datos_apoyo[fila_titulos+1:], columns=encabezados_limpios)

                    st.success("🛰️ Enlace Satelital Establecido. Pase al Módulo de Validación.")
                    
                    lista_pistas = []
                    for f in f_pistas:
                        nombre_archivo = f.name.lower()
                        bytes_f = io.BytesIO(f.getvalue())
                        dict_p = {}
                        
                        # 🛡️ REGLA DE ORO: Detectar y esquivar pestañas ocultas
                        if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xlsm'):
                            wb_temp = openpyxl.load_workbook(bytes_f, read_only=True)
                            # Extraemos SOLO los nombres de las hojas que están visibles
                            hojas_visibles = [ws.title for ws in wb_temp.worksheets if ws.sheet_state == 'visible']
                            bytes_f.seek(0) # Reiniciamos el lector
                            
                            if hojas_visibles:
                                dict_p = pd.read_excel(bytes_f, sheet_name=hojas_visibles, header=None)
                        
                        elif nombre_archivo.endswith('.xls'):
                            dict_p = pd.read_excel(bytes_f, sheet_name=None, header=None)
                        else:
                            try:
                                dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', header=None)}
                            except:
                                bytes_f.seek(0)
                                dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', encoding='latin1', header=None)}
                            
                        # Procesamos solo lo que pasó el filtro
                        for n, df in dict_p.items():
                            df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                            filas_c = df[df.astype(str).apply(lambda x: x.str.contains('COCTEL', case=False, na=False)).any(axis=1)].index.tolist()
                            for i, f_idx in enumerate(filas_c):
                                f_data = df.iloc[f_idx].dropna().tolist()
                                coctel = f_data[1] if len(f_data) > 1 else "S/N"
                                lim = filas_c[i+1] if i+1 < len(filas_c) else len(df)
                                segment = df.iloc[f_idx:lim]
                                idx_fin = segment[segment.astype(str).apply(lambda x: x.str.contains('FINCAS', case=False, na=False)).any(axis=1)].index
                                if not idx_fin.empty:
                                    f_h = idx_fin[0]
                                    c_idx = (df.iloc[f_h].astype(str).str.contains('FINCAS', case=False)).values.argmax()
                                    # --- 🛰️ NUEVO ESCÁNER DE BARRIDO MULTI-FINCA ---
                                    # --- 🛰️ NUEVO ESCÁNER MAESTRO: ANCLAJE POR PEDIDO SAP ---
                                    for r in range(f_h + 1, lim):
                                        # BLINDAJE: Convertimos cada celda a texto puro de Python para evitar choques con números/vacíos
                                        fila_textos = [str(x).strip() for x in df.iloc[r].tolist()]
                                        
                                        # 1. Freno de emergencia: Si dice TOTAL, cerramos el bloque
                                        if any("TOTAL" in celda.upper() for celda in fila_textos):
                                            break
                                            
                                        # 2. Radar de Pedido SAP: Buscamos de derecha a izquierda un número largo
                                        pedido_sap = ""
                                        for celda in reversed(fila_textos):
                                            # Buscamos números de mínimo 8 dígitos (Ej: 170036035)
                                            if celda.isdigit() and len(celda) >= 8: 
                                                pedido_sap = celda
                                                break
                                                
                                        # 3. Captura Confirmada: Si hay Pedido SAP, la fila tiene datos reales
                                        if pedido_sap:
                                            # Atrapamos la finca en su columna original o la de al lado
                                            fv = str(df.iloc[r, c_idx]).strip()
                                            if fv.lower() in ['nan', '', 'none', 'nat'] and (c_idx + 1) < len(df.columns):
                                                fv = str(df.iloc[r, c_idx + 1]).strip()
                                                
                                            if fv.lower() in ['nan', '', 'none', 'nat']:
                                                fv = "FINCA_SIN_NOMBRE" # Seguro de vida
                                                
                                            datos_fila = df.iloc[r].to_dict()
                                            
                                            lista_pistas.append({
                                                "ORIGEN": f"{f.name} | {n}", 
                                                "COCTEL": coctel, 
                                                "FINCA_INFORME": fv, 
                                                "PEDIDO_SAP": pedido_sap, # Guardamos el ancla
                                                "DATOS_FILA": datos_fila
                                            })
                                        st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                    st.balloons()
                except Exception as e: 
                    st.error(f"🚨 Error: {e}")

# =====================================================================
# ⚙️ 3. VALIDACIÓN DE MISIÓN (NÚCLEO FACTURACIÓN + SIMULADOR)
# =====================================================================
elif menu == "⚙️ 3. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>Núcleo de Validación y Facturación</h1>", unsafe_allow_html=True)
    
    # -----------------------------------------------------------------
    # 🔮 MODO SIMULADOR (MEGAZORD)
    # -----------------------------------------------------------------
    modo_simulacro = st.toggle("🔮 ACTIVAR MODO SIMULADOR (Modo Construcción de Matriz)")

    if modo_simulacro:
        st.info("💡 MODO CLON: Réplica exacta del Módulo de Validación con Cerebro Dinámico.")
        
        # --- 📡 1. CONEXIÓN A LA BÓVEDA ---
        if 'df_cfg' not in st.session_state or 'df_recetas' not in st.session_state or 'df_vd' not in st.session_state or 'df_t2' not in st.session_state:
            st.warning("⚠️ Bóveda Vacía. Conecte su Drive para cargar las matrices base.")
            url_drive = st.text_input("🔗 Pegue el Link de Google Drive (Google Sheets):", key="sim_drive")
            if url_drive:
                try:
                    import requests, io
                    file_id = url_drive.split('/d/')[1].split('/')[0] if '/d/' in url_drive else None
                    if file_id:
                        dl_url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx' if 'spreadsheets' in url_drive else f'https://drive.google.com/uc?export=download&id={file_id}'
                        with st.spinner("📥 Descargando matrices y TABLA 2..."):
                            resp = requests.get(dl_url, timeout=30)
                            if resp.status_code == 200:
                                xls = pd.ExcelFile(io.BytesIO(resp.content))
                                st.session_state['df_cfg'] = pd.read_excel(xls, sheet_name="Configuración")
                                st.session_state['df_recetas'] = pd.read_excel(xls, sheet_name="DD_Mesclas")
                                st.session_state['df_vd'] = pd.read_excel(xls, sheet_name="Validación Dosis")
                                
                                hojas = xls.sheet_names
                                nombre_tabla2 = "TABLA 2" if "TABLA 2" in hojas else hojas[1]
                                st.session_state['df_t2'] = pd.read_excel(xls, sheet_name=nombre_tabla2)
                                
                                st.success("✅ Matrices cargadas y listas.")
                                st.rerun()
                            else:
                                st.error(f"❌ Error de descarga: {resp.status_code}")
                    else:
                        st.error("❌ Link inválido.")
                except Exception as e:
                    st.error(f"🚨 Error: {e}")
            st.stop()

        df_cfg = st.session_state['df_cfg']
        df_recetas = st.session_state['df_recetas']
        df_vd = st.session_state['df_vd']
        df_t2 = st.session_state['df_t2']

        # --- 📡 2. EXTRACCIÓN ROBUSTA DE TOPES ---
        pistas_con_tope = []
        try:
            filas_a_revisar = [[str(c).upper().strip() for c in df_vd.columns]]
            for i in range(min(10, len(df_vd))): filas_a_revisar.append([str(x).upper().strip() for x in df_vd.iloc[i]])
            
            p_idx, t_idx, pr_idx = -1, -1, -1
            for idx_fila, row_vals in enumerate(filas_a_revisar):
                for i, val in enumerate(row_vals):
                    if val.startswith('TOPE'):
                        t_idx = i
                        for k in range(max(0, i-3), i):
                            if row_vals[k].startswith('PISTA'): p_idx = k
                            if 'PRECIO' in row_vals[k]: pr_idx = k
                if p_idx != -1 and t_idx != -1: break
                    
            if p_idx != -1 and t_idx != -1:
                for j in range(0, len(df_vd)):
                    p_name = str(df_vd.iloc[j, p_idx]).strip()
                    if p_name in ['NAN', 'NONE', ''] or pd.isna(df_vd.iloc[j, p_idx]): continue
                    p_tope = str(df_vd.iloc[j, t_idx]).strip()
                    if p_tope in ['NAN', 'NONE', '']: continue
                    p_precio = pd.to_numeric(df_vd.iloc[j, pr_idx], errors='coerce') if pr_idx != -1 else 0
                    if pd.isna(p_precio): p_precio = 0
                    texto_tope = f"{p_name} - {p_tope} (${p_precio:,.0f})".replace(',', '.')
                    if texto_tope not in pistas_con_tope: pistas_con_tope.append(texto_tope)
        except: pass
        
        if not pistas_con_tope: 
            pistas_con_tope = ["PLUC - TOPE MAX GENERAL ($63.325)", "PLUC - TOPE SUR ($70.829)", "PLUC - TOPE PARCELA INTER < 20ha ($98.335)", "PORI - TOPE MAX GENERAL ($62.718)", "PORI - TOPE SUR ($70.829)", "PORI - TOPE PARCELA INTER < 20ha ($105.723)", "PDIV - PORCION TERRESTRE ($8.504)", "TEHO - BASE ($0)", "LUCI - BASE ($0)"]

        # --- 🧠 3. CEREBRO DINÁMICO (TABLA 2) ---
        diccionario_fincas = {}
        lista_fincas = []
        try:
            for idx, row in df_t2.iterrows():
                f_name = str(row.iloc[0]).strip().upper()
                if f_name not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']:
                    p_tipo = str(row.iloc[5]).strip().upper() if len(row) > 5 else "TERCERO"
                    t_tipo = str(row.iloc[6]).strip().upper() if len(row) > 6 else ""
                    diccionario_fincas[f_name] = {"Productor": p_tipo, "Tope_Key": t_tipo}
                    if f_name not in lista_fincas: lista_fincas.append(f_name)
        except: pass
            
        if not lista_fincas: lista_fincas = ["NUEVO MUNDO"]
        lista_productores = ["SOCIO", "AGRICOLA", "AFILIADO", "TERCERO", "ORGANICO", "COOPERATIVA"]

        if 'finca_anterior' not in st.session_state:
            st.session_state.finca_anterior = lista_fincas[0]
            st.session_state.idx_prod = 3
            st.session_state.idx_tope = 0

        # --- 🎛️ 4. PANEL DE CONSTRUCCIÓN DINÁMICO ---
        st.markdown("#### 📝 Parámetros de la Operación")
        cs1, cs2, cs3, cs4 = st.columns(4)
        coctel_sim = cs1.text_input("🧪 Cóctel (Ej: IN6 ZN)", value="IN6")
        ha_sim = cs2.number_input("🚜 Hectáreas", min_value=1.0, value=143.0)
        finca_sim = cs3.selectbox("🏡 Finca", lista_fincas)
        
        if finca_sim != st.session_state.finca_anterior:
            datos = diccionario_fincas.get(finca_sim, {})
            if datos.get("Productor") in lista_productores: st.session_state.idx_prod = lista_productores.index(datos.get("Productor"))
            st.session_state.idx_tope = 0
            tope_k = datos.get("Tope_Key", "")
            if tope_k:
                for i, p_t in enumerate(pistas_con_tope):
                    if tope_k in p_t: st.session_state.idx_tope = i; break
            st.session_state.finca_anterior = finca_sim
            st.rerun()

        tipo_prod_sim = cs4.selectbox("🧑‍🌾 Productor (Márgenes)", lista_productores, index=st.session_state.idx_prod)
        
        st.markdown("<br>", unsafe_allow_html=True) 
        cs5, cs6, cs7, cs8 = st.columns(4)
        vuelo_sim = cs5.selectbox("🚁 Equipo", ["AVIÓN", "DRONE"])
        pista_sim = cs6.selectbox("🛣️ Pista y Tope", pistas_con_tope, index=st.session_state.idx_tope)
        horometro_sim = cs7.number_input("⏱️ Horómetro", min_value=0.01, value=3.30, step=0.1)
        dias_ciclo_sim = cs8.number_input("📅 Días Ciclo", min_value=0, value=14, step=1)
        
        recargo_sim = st.number_input("⚠️ Recargo ($/Ha)", min_value=0.0, value=5000.0, step=1000.0)

        if st.button("🚀 Construir Matriz MEGAZORD"):
            import re
            try:
                # 🎯 INTELIGENCIA DE MÁRGENES (Ajustado según Tabla de Configuración Oficial)
                if tipo_prod_sim == "TERCERO": mult_m = 1.451; st_base = 1583.0; mult_v = 1.451
                elif tipo_prod_sim == "AFILIADO": mult_m = 1.164; st_base = 1510.0; mult_v = 1.164
                elif tipo_prod_sim == "COOPERATIVA": mult_m = 1.112; st_base = 1510.0; mult_v = 1.164
                elif tipo_prod_sim == "ORGANICO": mult_m = 1.011; st_base = 1337.0; mult_v = 1.011
                else: mult_m = 1.112; st_base = 1337.0; mult_v = 1.112
                
                tarifa_vuelo_base = 4606562.0 

                val_tope = 0.0
                match = re.search(r'\(\$([\d\.]+)\)', pista_sim)
                if match: val_tope = float(match.group(1).replace('.', ''))

                # 🎯 INTELIGENCIA DE DRONES EN EL SIMULADOR SEGÚN PISTA
                if vuelo_sim == "DRONE": 
                    if "PLUC" in pista_sim: base_dron = 84428     # DATAROT
                    elif "PDIV" in pista_sim: base_dron = 76916   # NORTE
                    else: base_dron = 72600                       # AVIL / GENESYS
                    
                    unitario_vuelo = base_dron * mult_v
                else:
                    costo_bruto = (tarifa_vuelo_base * horometro_sim) / ha_sim if ha_sim > 0 else 0
                    if val_tope > 0: costo_bruto = min(costo_bruto, val_tope)
                    unitario_vuelo = costo_bruto * mult_v

                subtotal_vuelo = round(unitario_vuelo, 0) * ha_sim
                subtotal_st = round(st_base, 0) * dias_ciclo_sim * ha_sim

                coctel_u = coctel_sim.upper().strip()
                partes = coctel_u.split(" ")
                base_c = partes[0]
                sigla_f = partes[1] if len(partes) > 1 else ""

                receta_c = df_recetas[df_recetas.iloc[:,0].astype(str).str.upper() == base_c]
                prods_f = []
                for idx, row in receta_c.iterrows():
                    p = str(row.iloc[1]).upper().strip()
                    d = pd.to_numeric(row.iloc[2], errors='coerce')
                    if pd.notna(d) and d > 0 and p not in ['NAN', '']: prods_f.append({"PRODUCTO": p, "DOSIS": d})

                if sigla_f:
                    if "ZN" in sigla_f: prods_f.append({"PRODUCTO": "ZINTRAC X LITRO SV", "DOSIS": 0.5})
                    elif "BT" in sigla_f: prods_f.append({"PRODUCTO": "BANATREL SC", "DOSIS": 0.5})

                for item in prods_f:
                    if "ACONDICIONADOR" in item["PRODUCTO"]: item["DOSIS"] = 0.06 if ("ZN" in coctel_u or "BT" in coctel_u) else 0.02
                    elif "IMBIOSIL" in item["PRODUCTO"].replace(" ","") or "INBIOMAG" in item["PRODUCTO"]: item["DOSIS"] = 1.0 if sigla_f else 1.5

                tabla_visual = []
                mezcla_total = 0
                
                c_p_i, c_c_i = 8, 9 
                for i in range(5):
                    r_c = df_cfg.iloc[i].astype(str).str.upper().tolist()
                    if 'PRODUCTO' in r_c and 'COSTO' in r_c: c_p_i, c_c_i = r_c.index('PRODUCTO'), r_c.index('COSTO'); break

                for item in prods_f:
                    p, d = item["PRODUCTO"], item["DOSIS"]
                    mask = df_cfg.iloc[:, c_p_i].astype(str).str.upper().str.strip() == p
                    if mask.any():
                        p_b = pd.to_numeric(df_cfg[mask].iloc[0, c_c_i], errors='coerce')
                        if pd.notna(p_b):
                            p_m = p_b * mult_m
                            c_t_p = round((d * ha_sim) * p_m, 0)
                            mezcla_total += c_t_p
                            tabla_visual.append({"PRODUCTO": p, "DOSIS": f"{d:.3f}", "X": "-", "P. UNIT.": f"$ {p_b:,.0f}".replace(",","."), "P. + MARGEN": f"$ {p_m:,.0f}".replace(",","."), "COSTO TOTAL": f"$ {c_t_p:,.0f}".replace(",",".")})
                    else:
                        tabla_visual.append({"PRODUCTO": f"⚠️ {p}", "DOSIS": f"{d:.3f}", "X": "-", "P. UNIT.": "$ 0", "P. + MARGEN": "$ 0", "COSTO TOTAL": "$ 0"})

                recargo_m = round(recargo_sim * mult_v, 0)
                valor_recargo_t = recargo_m * ha_sim
                total_finca = subtotal_vuelo + subtotal_st + mezcla_total + valor_recargo_t
                costo_ha = total_finca / ha_sim if ha_sim > 0 else 0

                st.markdown("---")
                st.markdown(f"### 📋 MATRIZ DE VALIDACIÓN: {finca_sim}")
                st.caption(f"🗓️ **Días Ciclo:** {dias_ciclo_sim} | 🚜 **Área:** {ha_sim} Ha | 🧪 **Cóctel:** {coctel_sim}")
                st.dataframe(pd.DataFrame(tabla_visual), use_container_width=True, hide_index=True) 
                
                st.markdown("<br>", unsafe_allow_html=True)
                r1, r2, r3, r4, r5 = st.columns(5)
                r1.metric("👨‍🔬 Serv. Tec", f"$ {subtotal_st:,.0f}".replace(",", "."))
                r2.metric("✈️ Vuelo", f"$ {subtotal_vuelo:,.0f}".replace(",", "."))
                r3.metric("🧪 Mezcla", f"$ {mezcla_total:,.0f}".replace(",", "."))
                r4.metric("⚠️ Recargo", f"$ {valor_recargo_t:,.0f}".replace(",", "."))
                r5.markdown(f"<div style='background-color:#0d1b2a; padding:10px; border-radius:5px; border:1px solid #00ff00; text-align:center;'><p style='margin:0; color:#00ff00; font-size:12px;'>💰 COSTO x HA</p><h4 style='margin:0; color:white;'>$ {costo_ha:,.0f}</h4></div>", unsafe_allow_html=True)
                
                st.markdown("---")
                st.markdown(f"<h2 style='text-align: center; color: #d4af37;'>🔥 TOTAL OPERACIÓN: $ {total_finca:,.0f}</h2>".replace(",", "."), unsafe_allow_html=True)
            except Exception as e: st.error(f"Error: {e}")
        st.stop() # DETIENE LA EJECUCIÓN AQUÍ SI EL MODO SIMULADOR ESTÁ ACTIVO

    # -----------------------------------------------------------------
    # -----------------------------------------------------------------
    # ⚙️ MÓDULO ORIGINAL DE FACTURACIÓN (SE EJECUTA SI EL TOGGLE ESTÁ APAGADO)
    # -----------------------------------------------------------------
    if 'df_pistas' not in st.session_state or 'df_apoyo' not in st.session_state:
        st.warning("🚨 Cargue los archivos en el Módulo 2 e inicie el procesamiento.")
    else:
        with st.container(border=True):
            st.markdown("### 📡 Panel de Operaciones")
        
        # --- 🛰️ NUEVO RADAR SAP ---
        c_vacio, c_radar = st.columns([2, 2])
        pedido_sap = c_radar.text_input("📦 Buscar por N° Pedido SAP (Opcional):", key="buscar_sap_mod3", placeholder="Ej: 170036035")

        finca_sap = ""
        st.session_state['ha_radar_sap'] = 0.0  # Guardamos las Ha en memoria

        if pedido_sap and 'df_pedidos' in st.session_state:
            df_p = st.session_state['df_pedidos']
            match_sap = df_p[df_p.astype(str).apply(lambda x: x.str.contains(str(pedido_sap).strip())).any(axis=1)]
            
            if not match_sap.empty:
                try:
                    # 🎯 RECONOCIMIENTO DE COLUMNAS EXACTAS
                    col_finca = [c for c in df_p.columns if 'FINCA' in str(c).upper() or 'CLIENTE' in str(c).upper()][0]
                    col_ha = [c for c in df_p.columns if 'CANT' in str(c).upper() or 'HECT' in str(c).upper()][0]
                    col_mat = [c for c in df_p.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper()][0]
                    
                    finca_sap = str(match_sap.iloc[0][col_finca]).strip().upper()
                    
                    # 🎯 REGLA DE ORO 459: Francotirador directo a la columna Material
                    ha_correcta = 0.0
                    for _, fila_ped in match_sap.iterrows():
                        valor_material = str(fila_ped[col_mat]).strip()
                        if valor_material == "459" or valor_material.split(".")[0] == "459": 
                            ha_correcta = extraer_numero(fila_ped[col_ha])
                            break
                    
                    if ha_correcta > 0:
                        st.session_state['ha_radar_sap'] = ha_correcta
                    else:
                        st.session_state['ha_radar_sap'] = extraer_numero(match_sap.iloc[0][col_ha])
                    
                    st.success(f"✅ **SAP CONFIRMADO:** {finca_sap} | {st.session_state['ha_radar_sap']} Ha")
                except:
                    pass

        c0, c1, c2 = st.columns([1, 2, 2])
        fecha_operacion = c0.date_input("📅 Fecha de Vuelo", format="DD/MM/YYYY", key="fecha_vuelo_master")
        
        df_t2 = st.session_state.get('df_config', pd.DataFrame())
        lista_fincas = sorted(df_t2.iloc[:, 0].dropna().unique().tolist()) if not df_t2.empty else []
        opciones_finca = ["---"] + lista_fincas
        
        # 🎯 Inteligencia de auto-selección de Finca
        idx_finca = 0
        if finca_sap:
            for i, f in enumerate(opciones_finca):
                if f.upper() in finca_sap or finca_sap in f.upper():
                    idx_finca = i
                    break

        finca_sel = c1.selectbox("📍 Seleccione Finca:", opciones_finca, index=idx_finca)
        
        vuelos_informe = st.session_state.get('df_pistas', pd.DataFrame())
        lista_origenes = vuelos_informe['ORIGEN'].unique().tolist() if not vuelos_informe.empty else []
        vuelo_ref = c2.selectbox("📄 Referencia Pedido/Informe:", ["---"] + lista_origenes)

        if finca_sel == "---" or vuelo_ref == "---":
            st.info("⚠️ Seleccione Finca y Pedido para rugir motores.")
            st.stop()

        # --- 🛰️ EXTRACCIÓN DE INTELIGENCIA DE COSTOS ---
        mult_material = 1.112; tarifa_serv_tec_base = 1337.0; mult_avion_base = 1.112
        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        import re 
        finca_limpia = re.sub(r'\s+', ' ', str(finca_sel)).strip().upper()

        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        
        if not df_t2.empty:
            match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia]
            if not match_t2.empty:
                fila_t2 = match_t2.iloc[0]
                tipo_productor = str(fila_t2.iloc[5]).strip().upper()
                tipo_de_tope_finca = str(fila_t2.iloc[6]).strip().upper()
        
        if not df_cfg.empty:
            match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
            if not match_cfg.empty:
                fila_c = match_cfg.iloc[0]
                mult_material = extraer_numero(fila_c.iloc[3])
                tarifa_serv_tec_base = extraer_numero(fila_c.iloc[4])
                mult_avion_base = extraer_numero(fila_c.iloc[6])
                
        dias_ciclo_calc = 0
        if not df_apoyo.empty:
            col_finca = [c for c in df_apoyo.columns if 'FINCA' in str(c).upper()]
            col_fecha = [c for c in df_apoyo.columns if 'FECHA' in str(c).upper()]
            if col_finca and col_fecha:
                mask_finca = df_apoyo[col_finca[0]].apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia
                hist_finca = df_apoyo[mask_finca].copy()
                if not hist_finca.empty:
                    hist_finca['FECHA_DT'] = hist_finca[col_fecha[0]].apply(procesar_fecha_pesada)
                    hist_finca = hist_finca.dropna(subset=['FECHA_DT'])
                    if not hist_finca.empty:
                        fecha_ref = pd.to_datetime(fecha_operacion)
                        vuelos_anteriores = hist_finca[hist_finca['FECHA_DT'] < fecha_ref]
                        if not vuelos_anteriores.empty:
                            dias_ciclo_calc = (fecha_ref - vuelos_anteriores['FECHA_DT'].max()).days

        datos_vuelo = vuelos_informe[vuelos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo.get('DATOS_FILA', {})
        
        num_pedido = "S/N"
        if pedido_sap and len(str(pedido_sap)) >= 7:
            num_pedido = str(pedido_sap).strip()
        elif datos_vuelo.get('PEDIDO_SAP') and str(datos_vuelo.get('PEDIDO_SAP')).strip() != "":
            num_pedido = str(datos_vuelo.get('PEDIDO_SAP')).strip()
        else:
            for idx in range(18, 40):
                val_celda = str(datos_raw.get(idx, "")).split('.')[0].strip()
                if val_celda.isdigit() and len(val_celda) >= 7:
                    num_pedido = val_celda
                    break
        
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
                for _, r_p in match_ped.iterrows():
                    if len(r_p) >= 7 and "459" in str(r_p.iloc[5]):
                        ha_dosis_detectada = extraer_numero(r_p.iloc[6])
                        break
        
        ha_cobro_detectada = extraer_numero(datos_raw.get(8, 0))
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        casilla_key = f"{finca_sel}_{vuelo_ref}_{fecha_operacion}"
        
        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros Base e Inteligencia de Ciclos")
            c_sup1, c_sup2 = st.columns([3, 1])
            c_sup1.info(f"🧑‍🌾 Productor: **{tipo_productor}** | 🛣️ Tope: **{tipo_de_tope_finca}**")
            
            mision_solo_dron = c_sup2.toggle("🚁 MISIÓN 100% DRON", value=False, key=f"dron_toggle_{casilla_key}")
            
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            r1c1.number_input("📅 Ciclo (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key=f"ds_{casilla_key}")
            d_ciclo_factura = r1c2.number_input("⏳ Ciclo (COBRO)", value=int(dias_ciclo_calc), step=1, key=f"df_{casilla_key}")
            
            ha_sugerida = float(st.session_state.get('ha_radar_sap', 0.0))
            if ha_sugerida == 0.0: ha_sugerida = float(ha_dosis_detectada)
                
            ha_dosis_final = r1c3.number_input("🧪 Ha Dosis (Total 459)", value=ha_sugerida, key=f"had_{casilla_key}")
            
            multi_aviones = r1c4.toggle("✈️ Recargo Coord. Multi-Avión", value=False, key=f"ma_{casilla_key}")
            mult_avion_final = mult_avion_base + 0.1 if multi_aviones else mult_avion_base

            recargo_final = 0.0
            pista_sel = "PLUC"
            if not mision_solo_dron:
                st.markdown("##### 🛣️ Parámetros Terrestres (Aviones)")
                r2c1, r2c2, r2c3 = st.columns(3)
                pista_sugerida = next((p for p in lista_pistas_validas if p in pista_detectada), "PLUC")
                pista_sel = r2c1.selectbox("Pista Base", lista_pistas_validas, index=lista_pistas_validas.index(pista_sugerida), key=f"pi_{casilla_key}")
                
                opciones_rec = ["0 (Sin Recargo)", "8504 (Porción PDIV)", "45000 (Recargo T. General)", "Otro Valor Manual..."]
                idx_recargo = 1 if pista_sel == "PDIV" else 0 
                recargo_lista = r2c2.selectbox("🚛 Recargo Terrestre:", opciones_rec, index=idx_recargo, key=f"rl_{casilla_key}")
                if recargo_lista == "Otro Valor Manual...":
                    recargo_final = r2c3.number_input("✍️ Digite Recargo ($)", value=0, step=1000, key=f"rm_{casilla_key}")
                else:
                    recargo_final = float(recargo_lista.split(" ")[0])

        dict_topes_pista = {"TOPE MAX GENERAL": {"PLUC": 63326, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325}, "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517}, "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}}
        val_tope = dict_topes_pista.get(tipo_de_tope_finca, {}).get(pista_sel, 999999)
        
        with st.container(border=True):
            st.markdown("#### ✈️ Hangar de Despliegue")
            costo_total_vuelos = 0.0
            costo_neto_vuelo_total = 0.0  
            total_ha_cobro_escuadron = 0.0
            horometro_final_avion = 0.0 

            if mision_solo_dron:
                st.success("🚁 Modo Dron Activo: Costos calculados sin recargos terrestres ni topes de pista.")
                try:
                    if "gcp_credentials" in st.secrets:
                        gc_vd = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc_vd = gspread.service_account(filename='credenciales.json')
                    boveda_vd = gc_vd.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    datos_vd = boveda_vd.worksheet("Validación Dosis").get_all_values()
                    df_flota = pd.DataFrame(datos_vd[2:], columns=datos_vd[1]) 
                    df_dr = df_flota[df_flota['Tarifa'].notna() & (df_flota['Tarifa'].astype(str).str.strip() != '')]
                    nombres_dr = df_dr['Tarifa'].astype(str).str.replace('TARIFA ', '', case=False).str.strip()
                    nombres_dr = nombres_dr.apply(lambda x: f"DRONE {x}" if "DRONE" not in x.upper() else x)
                    precios_dr = pd.to_numeric(df_dr['Valor ha/Dr'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)
                    dict_drones = dict(zip(nombres_dr, precios_dr))
                except Exception as e:
                    dict_drones = {"DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280}

                # Lee el total del piloto por defecto
                df_drones_def = pd.DataFrame([{"Drone": "DRONE DATAROT", "Hectáreas": float(ha_cobro_detectada)}])
                escuadron_drones = st.data_editor(df_drones_def, key=f"drones_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                    costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                    costo_total_vuelos += (tarifa_dron_neta * ha_dr) * mult_avion_final 

            else:
                c_av, c_dr = st.columns(2)
                try:
                    if "gcp_credentials" in st.secrets:
                        gc_vd = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc_vd = gspread.service_account(filename='credenciales.json')
                    boveda_vd = gc_vd.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    datos_vd = boveda_vd.worksheet("Validación Dosis").get_all_values()
                    df_flota = pd.DataFrame(datos_vd[2:], columns=datos_vd[1]) 
                    df_av = df_flota[df_flota['TIPO'].notna() & (df_flota['TIPO'].astype(str).str.strip() != '')]
                    dict_aviones = dict(zip(df_av['TIPO'].astype(str).str.strip(), pd.to_numeric(df_av['HORA'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)))
                    df_dr = df_flota[df_flota['Tarifa'].notna() & (df_flota['Tarifa'].astype(str).str.strip() != '')]
                    nombres_dr = df_dr['Tarifa'].astype(str).str.replace('TARIFA ', '', case=False).str.strip()
                    nombres_dr = nombres_dr.apply(lambda x: f"DRONE {x}" if "DRONE" not in x.upper() else x)
                    precios_dr = pd.to_numeric(df_dr['Valor ha/Dr'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)
                    dict_drones = dict(zip(nombres_dr, precios_dr))
                except Exception as e:
                    dict_aviones = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA 25": 3036525, "AIR TRACTOR": 4665109, "CESSNA ASA": 3666600, "CESSNA FUMIGARAY": 3065952}
                    dict_drones = {"DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280}

                with c_av: 
                    st.markdown("##### 🛩️ Base Aviones")
                    # Lee el total del piloto por defecto
                    df_aviones_def = pd.DataFrame([{"Avión": "CESSNA ASA", "Hectáreas": float(ha_cobro_detectada), "Horómetro": 1.00}])
                    opciones_av = list(dict_aviones.keys()) if 'dict_aviones' in locals() and dict_aviones else ["THRUS SR2", "PIPER PA 36-375"]
                    escuadron_aviones = st.data_editor(df_aviones_def, key=f"aviones_{casilla_key}", num_rows="dynamic", column_config={"Avión": st.column_config.SelectboxColumn("Modelo", options=opciones_av, required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f"), "Horómetro": st.column_config.NumberColumn("Horómetro", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)
                    
                with c_dr:
                    st.markdown("##### 🚁 Base Drones (Apoyo)")
                    df_drones_def = pd.DataFrame([{"Drone": None, "Hectáreas": 0.0}])
                    opciones_dr = list(dict_drones.keys()) if 'dict_drones' in locals() and dict_drones else ["DRONE DATAROT", "DRON GENESYS"]
                    escuadron_drones = st.data_editor(df_drones_def, key=f"drones_mix_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=opciones_dr), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)                
                
                # 🛡️ CÁLCULOS PROTEGIDOS E INTELIGENCIA NETA
                for index, row in escuadron_aviones.iterrows():
                    av_sel = row["Avión"]
                    try: ha_av = float(row.get("Hectáreas", 0)) if str(row.get("Hectáreas", 0)) not in ["None", "", "nan"] else 0.0
                    except: ha_av = 0.0
                        
                    try: horo = float(row.get("Horómetro", 0)) if str(row.get("Horómetro", 0)) not in ["None", "", "nan"] else 0.0
                    except: horo = 0.0
                    
                    if pd.isna(av_sel) or ha_av <= 0: continue
                    total_ha_cobro_escuadron += ha_av
                    horometro_final_avion += horo  
                    
                    tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av if ha_av > 0 else 0
                    tarifa_base_tope = tarifa_base_ha if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope)
                    
                    costo_neto_vuelo_total += (tarifa_base_tope * ha_av) 
                    tarifa_aplicada = tarifa_base_tope + recargo_final
                    costo_total_vuelos += (tarifa_aplicada * ha_av) * mult_avion_final 
                    
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                    costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                    costo_total_vuelos += (tarifa_dron_neta * ha_dr) * mult_avion_final
            
        st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Mezcla")
        # 🎯 PUENTE DE MANDO: Control maestro de pista ANTES de armar la matriz
        pistas_disponibles = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"]
        idx_pista = pistas_disponibles.index(pista_sel) if 'pista_sel' in locals() and pista_sel in pistas_disponibles else 0
        
        pista_sel = st.selectbox("📍 Seleccione la Pista para extraer Inventario de SAP:", pistas_disponibles, index=idx_pista, key="pista_matriz_maestra")
        
        st.markdown("---")
        costo_mezcla_total = 0.0

        if not match_ped.empty:
            idx_precio = -1; idx_lote = -1; idx_saldo = -1; idx_almacen = -1
            if not df_sab.empty:
                for j, col in enumerate(df_sab.columns):
                    col_str = str(col).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').strip()
                    
                    if ('MAYOR' in col_str or 'PRECIO' in col_str) and idx_precio == -1: idx_precio = j
                    if 'LOTE' in col_str and 'PROVEEDOR' not in col_str and idx_lote == -1: idx_lote = j
                    # 🎯 RADAR DE PISTA: Excluye la trampa de la columna "PB"
                    if ('ALMACEN' in col_str or 'PISTA' in col_str) and 'PB' not in col_str and idx_almacen == -1: idx_almacen = j
                    if ('LIBRE' in col_str or 'SALDO' in col_str) and 'VALOR' not in col_str and idx_saldo == -1: idx_saldo = j
                        
            sap_dict_pista = {}
            datos_extraidos_sap = []

            for _, fila_sap in match_ped.iterrows():
                col_mat = [c for c in fila_sap.index if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper() or 'COD' in str(c).upper()]
                if not col_mat: continue
                texto_material = str(fila_sap[col_mat[0]]).strip()
                if "459" in texto_material or "429" in texto_material: continue

                cod_item = texto_material.split('.')[0].lstrip('0')
                col_cant = [c for c in fila_sap.index if 'DOSIS' in str(c).upper() or 'CANT' in str(c).upper()]
                cant_total = extraer_numero(fila_sap[col_cant[0]]) if col_cant else 0.0
                dosis_pista = cant_total / ha_dosis_final if ha_dosis_final > 0 else 0.0

                nombre_p = f"Item {cod_item}"
                if not df_sab.empty:
                    match_sabana = df_sab[df_sab.iloc[:, 0].astype(str).str.strip() == cod_item]
                    if match_sabana.empty: match_sabana = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(cod_item, case=False, na=False)).any(axis=1)]
                    if not match_sabana.empty:
                        col_nombre_sab = [c for c in match_sabana.columns if 'TEXTO' in str(c).upper() or 'DESC' in str(c).upper()]
                        if col_nombre_sab: nombre_p = str(match_sabana.iloc[0][col_nombre_sab[0]]).upper()

                nombre_limpio = nombre_p.split('*')[0].strip().replace(" ", "")
                
                # 🧠 ACUMULADOR PARA LA IA: Sumamos la dosis para que adivine el Cóctel Correcto
                sap_dict_pista[nombre_limpio] = sap_dict_pista.get(nombre_limpio, 0.0) + dosis_pista
                
                # 📋 Mantenemos el registro individual para que la tabla muestre los lotes separados
                datos_extraidos_sap.append({"cod": cod_item, "nombre": nombre_p, "nombre_limpio": nombre_limpio, "cant_total": cant_total})
            dict_recetas = {}
            dict_lideres = {}
            dict_fertilizantes = {}

            if not df_mez.empty:
                for idx, row in df_mez.iterrows():
                    if len(row) > 3:
                        cid = str(row.iloc[0]).strip().upper()
                        p_tabla_clean = str(row.iloc[1]).strip().upper().replace(" ", "")
                        d_tabla = extraer_numero(row.iloc[2])
                        es_lider = str(row.iloc[3]).strip().upper() == "X"
                        if cid and cid != 'NAN' and p_tabla_clean:
                            if cid not in dict_recetas: dict_recetas[cid] = {}
                            dict_recetas[cid][p_tabla_clean] = d_tabla
                            if es_lider: dict_lideres[cid] = p_tabla_clean
                    if len(row) > 13:
                        fert_name = str(row.iloc[12]).strip().upper()
                        fert_sigla = str(row.iloc[13]).strip().upper()
                        if fert_name and fert_sigla and fert_name not in ["NAN", "FERTILIZANTES", ""]:
                            dict_fertilizantes[fert_name.replace(" ", "")] = fert_sigla

            coctel_base = "SIN COINCIDENCIA"
            dosis_oficiales_coctel = {}
            max_p = -999

            # 🎯 INTELIGENCIA DE PILOTO: Extraemos el cóctel exacto del informe (Ej: SKMN53+GL -> SKMN53)
            coctel_piloto_raw = str(datos_vuelo.get('COCTEL', '')).upper().strip()
            coctel_piloto_base = coctel_piloto_raw.replace("+", " ").replace("-", " ").split(" ")[0]

            for iter_id, receta in dict_recetas.items():
                es_valido = True
                puntaje = 0
                lider_db = dict_lideres.get(iter_id, "")
                match_lider = False
                if lider_db:
                    for k_sap in sap_dict_pista.keys():
                        if lider_db == k_sap or (len(k_sap)>=4 and lider_db in k_sap) or (len(lider_db)>=4 and k_sap in lider_db):
                            match_lider = True; break
                if match_lider: puntaje += 1000
                else: es_valido = False

                if es_valido:
                    # 🎯 BONO FRANCOTIRADOR: Si la base coincide con lo que anotó el piloto, gana automáticamente
                    if iter_id == coctel_piloto_base:
                        puntaje += 10000

                    for p_receta, d_esperada in receta.items():
                        match_receta = False
                        dose_matched = False
                        for k_sap, d_sap in sap_dict_pista.items():
                            if p_receta == k_sap or (len(k_sap)>=4 and p_receta in k_sap) or (len(p_receta)>=4 and k_sap in p_receta):
                                match_receta = True
                                # Ampliamos la tolerancia a 0.5 para que no descarte por decimales en SAP
                                if abs(d_sap - d_esperada) <= 0.5: dose_matched = True 
                                break
                        if match_receta: puntaje += 50 if dose_matched else 10
                        else: es_valido = False; break

                if es_valido and puntaje > max_p:
                    max_p = puntaje
                    coctel_base = iter_id
                    dosis_oficiales_coctel = receta.copy()
                    
            # --- FASE 2: BUSCAR EL FERTILIZANTE Y SU SIGLA ---
            sigla_fertilizante = ""
            for k_sap in sap_dict_pista.keys():
                for f_name, f_sigla in dict_fertilizantes.items():
                    if f_name == k_sap or (len(k_sap)>=4 and f_name in k_sap) or (len(f_name)>=4 and k_sap in f_name):
                        sigla_fertilizante = f" {f_sigla}" # 🎯 Aquí quitamos el "+" y dejamos un espacio
                        break
                if sigla_fertilizante: break

            coctel_ganador = coctel_base + sigla_fertilizante if coctel_base != "SIN COINCIDENCIA" else "SIN COINCIDENCIA"

            st.success(f"🤖 **MOTOR IA MAESTRO:** Cóctel Oficial: **{coctel_ganador}**")

            # --- 🔴 INYECCIÓN DE RAYOS X PARA DIAGNÓSTICO ---
            st.error(f"🔍 **RAYOS X DE COLUMNAS:** Pista Elegida: '{pista_sel}' | Lote: Columna {idx_lote} | Almacén: Columna {idx_almacen} | Saldo: Columna {idx_saldo}")
            
            matriz_datos = []
            for item_data in datos_extraidos_sap:
                # Quitamos ceros a la izquierda del Pedido (Ej: 000300054 -> 300054)
                cod_item = str(item_data['cod']).strip().upper().lstrip('0')
                nombre_p = item_data['nombre']
                nombre_limpio = item_data['nombre_limpio']
                cant_total_pedido = item_data['cant_total']

                costo_unit = 0.0; lote_sap = "SIN LOTE EN PISTA"; saldo_sap = 0.0

                if not df_sab.empty:
                    # 1. BÚSQUEDA BLINDADA (Quitando ceros a la izquierda también en la Sábana)
                    col_0_limpia = df_sab.iloc[:, 0].apply(lambda x: str(x).split('.')[0].strip().upper().lstrip('0') if str(x).lower() not in ['nan', 'none', '<na>', ''] else "")
                    match_sabana_global = df_sab[col_0_limpia == cod_item]
                    
                    # 2. SALVAVIDAS: Buscar por nombre si el código falla
                    if match_sabana_global.empty and nombre_limpio != "" and "ITEM" not in nombre_limpio:
                        match_sabana_global = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(nombre_limpio, case=False, na=False)).any(axis=1)]

                    if not match_sabana_global.empty:
                        # Extraer Precio
                        fila_precio = match_sabana_global.iloc[0]
                        if idx_precio != -1: 
                            costo_unit = extraer_numero(fila_precio.iloc[idx_precio])
                        if costo_unit == 0.0:
                            col_v = [c for c in fila_precio.index if 'VALOR' in str(c).upper() and 'LIBRE' in str(c).upper()]
                            col_c = [c for c in fila_precio.index if 'LIBRE' in str(c).upper() and 'VALOR' not in str(c).upper()]
                            if col_v and col_c:
                                v_t = extraer_numero(fila_precio[col_v[0]])
                                c_t = extraer_numero(fila_precio[col_c[0]])
                                if c_t > 0: costo_unit = v_t / c_t

                        # 3. FILTRAR POR PISTA SELECCIONADA
                        if idx_almacen != -1:
                            col_almacen = match_sabana_global.iloc[:, idx_almacen].astype(str).str.strip().str.upper()
                            match_pista = match_sabana_global[col_almacen.str.contains(str(pista_sel).strip().upper(), na=False)]
                        else:
                            match_pista = match_sabana_global[match_sabana_global.astype(str).apply(lambda x: x.str.strip().str.upper().str.contains(str(pista_sel).strip().upper(), na=False)).any(axis=1)]

                        # 4. EXTRAER LOTE Y SALDO (FIFO)
                        if not match_pista.empty:
                            try:
                                if idx_saldo != -1:
                                    match_pista['Temp_Sort'] = match_pista.iloc[:, idx_saldo].apply(extraer_numero)
                                    match_vivos = match_pista[match_pista['Temp_Sort'] > 0]
                                    match_pista = match_vivos.sort_values(by='Temp_Sort', ascending=True) if not match_vivos.empty else match_pista.sort_values(by='Temp_Sort', ascending=False)
                            except: pass
                            
                            fila_final = match_pista.iloc[0]
                            if idx_lote != -1: lote_sap = str(fila_final.iloc[idx_lote])
                            if idx_saldo != -1: saldo_sap = extraer_numero(fila_final.iloc[idx_saldo])

                # 🛡️ CÁLCULO DE DOSIS
                total_sap_producto = sum(item['cant_total'] for item in datos_extraidos_sap if item['cod'] == item_data['cod'])
                dosis_teorica = None
                for p_receta, d_oficial in dosis_oficiales_coctel.items():
                    if p_receta == nombre_limpio or (len(nombre_limpio)>=4 and p_receta in nombre_limpio) or (len(p_receta)>=4 and nombre_limpio in p_receta):
                        dosis_teorica = d_oficial; break

                if "ACONDICIONADOR" in nombre_limpio:
                    dosis_teorica = 0.06 if ("ZN" in coctel_ganador or "BT" in coctel_ganador) else 0.02
                elif "IMBIOSIL" in nombre_limpio.replace(" ", "") or "INBIOMAG" in nombre_limpio:
                    dosis_teorica = 1.5 if coctel_ganador.startswith("IN") else 1.0
                
                if dosis_teorica is None:
                    dosis_teorica = total_sap_producto / ha_dosis_final if ha_dosis_final > 0 else 0.0
                    
                costo_margen = round(costo_unit * mult_material, 0)

                # EMPAQUETADO FINAL
                matriz_datos.append({
                    "A: Producto": nombre_p,
                    "B: Dosis/Ha (SAP)": round(dosis_teorica, 3),
                    "C: X (Extra %)": 0.0,
                    "D: Dosis Total (Sistema)": 0.0,
                    "E: Costo Unit (+Margen)": costo_margen,
                    "G: Lotes (SAP)": lote_sap,
                    "H: Saldo Real SAP": round(saldo_sap, 3),
                    "I: Sugerido SAP (Total)": round(cant_total_pedido, 3)
                })

            df_matriz = pd.DataFrame(matriz_datos)
                
            if 'editor_valid' in st.session_state:
                ediciones = st.session_state['editor_valid'].get('edited_rows', {})
                for row_idx, edit_dict in ediciones.items():
                    if "B: Dosis/Ha (SAP)" in edit_dict: df_matriz.at[row_idx, "B: Dosis/Ha (SAP)"] = edit_dict["B: Dosis/Ha (SAP)"]
                    if "C: X (Extra %)" in edit_dict: df_matriz.at[row_idx, "C: X (Extra %)"] = edit_dict["C: X (Extra %)"]

            df_matriz["B_Val"] = df_matriz["B: Dosis/Ha (SAP)"].fillna(0.0)
            df_matriz["C_Val"] = df_matriz["C: X (Extra %)"].fillna(0.0)
            df_matriz["D: Dosis Total (Sistema)"] = (df_matriz["B_Val"] * (1 + df_matriz["C_Val"]/100) * ha_dosis_final).round(3)

            # --- 🧪 CÁLCULO DE MEZCLA CON REDONDEO ARITMÉTICO (ESTILO SAP) ---
            import math
            df_matriz["Total_Fila"] = (df_matriz["I: Sugerido SAP (Total)"] * df_matriz["E: Costo Unit (+Margen)"])
            costo_mezcla_total = df_matriz["Total_Fila"].apply(lambda x: math.floor(x + 0.5)).sum()
            
            df_matriz = df_matriz.drop(columns=["B_Val", "C_Val", "Total_Fila"])

            # 🔥 RADAR DE COLOR TÉRMICO (FORMATO CONDICIONAL INTELIGENTE)
            def colorear_matriz(row):
                # El sistema suma todas las líneas de este producto para validar el global
                global_sap = df_matriz[df_matriz["A: Producto"] == row["A: Producto"]]["I: Sugerido SAP (Total)"].sum()
                ideal_sistema = row["D: Dosis Total (Sistema)"]
                
                # Calculamos la diferencia matemática
                diferencia = abs(global_sap - ideal_sistema)
                
                # 🎨 Pintura Táctica según el desfase
                if diferencia <= 0.5:
                    color = 'background-color: #d4edda; color: #155724;' # 🟢 Verde: Cuadrado Perfecto
                elif diferencia <= 5.0:
                    color = 'background-color: #fff3cd; color: #856404;' # 🟡 Amarillo: Leve Desfase
                elif diferencia <= 20.0:
                    color = 'background-color: #f8d7da; color: #721c24;' # 🔴 Rojo Claro: Alerta Media
                else:
                    color = 'background-color: #8b0000; color: white; font-weight: bold;' # 🚨 Rojo Sangre: Crítico
                    
                return [color] * len(row)

            # Bañamos la matriz con la inteligencia de colores
            df_pintado = df_matriz.style.apply(colorear_matriz, axis=1)

            edited_df = st.data_editor(
                df_pintado, key='editor_valid',
                column_config={
                    "B: Dosis/Ha (SAP)": st.column_config.NumberColumn("Dosis/Ha", min_value=0.000, format="%.3f"),
                    "C: X (Extra %)": st.column_config.NumberColumn("Extra %", min_value=0.000, format="%.3f"),
                    "D: Dosis Total (Sistema)": st.column_config.NumberColumn("Dosis Ideal", format="%.3f"),
                    "E: Costo Unit (+Margen)": st.column_config.NumberColumn("Costo Unit (COP)", format="%.0f"),
                    "H: Saldo Real SAP": st.column_config.NumberColumn("Saldo SAP", format="%.3f"),
                    "I: Sugerido SAP (Total)": st.column_config.NumberColumn("Sugerido SAP (Total)", format="%.3f"),
                },
                disabled=["A: Producto", "D: Dosis Total (Sistema)", "E: Costo Unit (+Margen)", "G: Lotes (SAP)", "H: Saldo Real SAP", "I: Sugerido SAP (Total)"],
                use_container_width=True, hide_index=True
            )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("##### 📋 Copia Rápida para SAP (Costo Unitario)")
  
            
            costos_limpios = df_matriz['E: Costo Unit (+Margen)'].fillna(0).astype(int).astype(str).tolist()
            texto_para_copiar = "\n".join(costos_limpios)
            st.code(texto_para_copiar, language="text")

        else:
            st.warning("🚨 No se encontró un pedido válido para la matriz de químicos.")
            costo_mezcla_total = 0.0

        st.markdown("---")
        st.markdown("### 💰 Liquidación Final (Bóveda SAP)")
        
        # =======================================================
        # --- 1. CÁLCULOS CON PRECISIÓN ARITMÉTICA (ESTILO SAP) ---
        # =======================================================
        import math

        # Función de redondeo contable (Arriba de .5 siempre sube)
        def sap_round(n):
            return math.floor(n + 0.5)

        # Cálculos de unitarios
        unitario_st_bruto = d_ciclo_factura * tarifa_serv_tec_base
        unitario_vuelo_bruto = costo_total_vuelos / total_ha_cobro_escuadron if total_ha_cobro_escuadron > 0 else 0
        
        # Aplicamos el redondeo de SAP a los unitarios
        unitario_st = sap_round(unitario_st_bruto)
        unitario_vuelo = sap_round(unitario_vuelo_bruto)
        
        # Subtotales redondeados línea por línea (Como hace SAP)
        subtotal_st_finca = sap_round(unitario_st * ha_dosis_final)
        subtotal_vuelo_finca = sap_round(unitario_vuelo * ha_dosis_final)
        
        # 🔥 GRAN TOTAL: Suma de subtotales ya redondeados aritméticamente
        gran_total = costo_mezcla_total + subtotal_vuelo_finca + subtotal_st_finca
        
        # Costo por Hectárea derivado del total final
        costo_por_ha = sap_round(gran_total / ha_dosis_final) if ha_dosis_final > 0 else 0

        # --- 🛰️ CÁLCULO DE REFERENCIA PARA GERENCIA ---
        # Sacamos el precio base de la hora del primer avión del escuadrón
        precio_hora_referencia = 0
        if not mision_solo_dron:
            try:
                if not escuadron_aviones.empty:
                    avion_principal = escuadron_aviones.iloc[0]['Avión']
                    precio_hora_referencia = dict_aviones.get(avion_principal, 0)
            except:
                precio_hora_referencia = 0

        # 🚁 NUEVO: Rescatar el precio base del Dron seleccionado
        precio_dron_referencia = 0
        try:
            # Buscamos en la tabla de drones (sea misión mixta o solo dron)
            df_busqueda_dron = escuadron_drones if mision_solo_dron else escuadron_drones
            if not df_busqueda_dron.empty:
                dron_principal = df_busqueda_dron.iloc[0]['Drone']
                if pd.notna(dron_principal) and str(dron_principal).strip() not in ["None", ""]:
                    precio_dron_referencia = dict_drones.get(dron_principal, 0)
        except:
            precio_dron_referencia = 0

        # --- 2. MÉTRICAS VISUALES (HTML PERSONALIZADO ALTO CONTRASTE) ---
        st.markdown("---")
        st.markdown("<br>", unsafe_allow_html=True)
        
        m1, m2, m3, m4, m5 = st.columns(5)
        
        def mini_metric(icono, titulo, valor):
            return f"""
            <div style='background-color: #0d1b2a; padding: 10px; border-radius: 8px; border-left: 4px solid #d4af37; box-shadow: 2px 2px 5px rgba(0,0,0,0.1); margin-bottom: 10px;'>
                <p style='margin:0; font-size: 0.75rem; color: #d4af37; text-transform: uppercase;'>{icono} {titulo}</p>
                <p style='margin:0; font-size: 1.15rem; font-weight: bold; color: white;'>{valor}</p>
            </div>
            """

        with m1: st.markdown(mini_metric("🚜", "Hectáreas", f"{ha_dosis_final:.2f} Ha"), unsafe_allow_html=True)
        # 🎯 ZONA DE CONTROL TOPE (M2): Nombre del tope arriba, Valor del tope abajo
        with m2: 
            # 🕵️‍♂️ Radar de actividad real: ¿Voló el avión o todo el trabajo lo hizo el Dron?
            ha_avion_real = 0
            try:
                if not mision_solo_dron and not escuadron_aviones.empty:
                    ha_avion_real = float(escuadron_aviones['Hectáreas'].sum())
            except:
                ha_avion_real = 0

            # Si el switch es 100% Dron, o si el avión tiene 0 Ha y hay un Dron asignado, domina el Dron
            es_dron_dominante = mision_solo_dron or (ha_avion_real == 0 and precio_dron_referencia > 0)

            # Pintamos la tarjeta de la pista o asumimos DRON
            st.markdown(mini_metric("🛣️", "Pista", tipo_de_tope_finca if not es_dron_dominante else "DRON"), unsafe_allow_html=True)
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True) # Espacio separador
            
            # 🛡️ Inteligencia para mostrar el valor del tope o la tarifa del Dron
            if es_dron_dominante:
                texto_valor_tope = f"$ {fmt_sap(precio_dron_referencia)}"
            elif val_tope == 999999 or val_tope == 0:
                texto_valor_tope = "Sin Tope"
            else:
                texto_valor_tope = f"$ {fmt_sap(val_tope)}"
                
            st.markdown(mini_metric("🚧", "Valor Tope", texto_valor_tope), unsafe_allow_html=True)
        with m3: st.markdown(mini_metric("👨‍🔬", "Tarifa ST", f"$ {fmt_sap(tarifa_serv_tec_base)}"), unsafe_allow_html=True)
        with m4: st.markdown(mini_metric("✈️", "Mult.", f"x {mult_avion_final}"), unsafe_allow_html=True)
        
        # 🎯 ZONA DE CONTROL GLOBAL (M5): Avión arriba, Dron abajo
        with m5: 
            st.markdown(mini_metric("⏱️", "Precio Hora", f"$ {fmt_sap(precio_hora_referencia)}"), unsafe_allow_html=True)
            st.markdown(mini_metric("🚁", "Tarifa Dron", f"$ {fmt_sap(precio_dron_referencia)}"), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Cajas de Copia para Digitación en SAP")
        
        c_sap1, c_sap2, c_sap3, c_sap4 = st.columns(4)
        with c_sap1: 
            st.caption("👨‍🔬 UNITARIO ST (459)")
            st.code(fmt_sap(unitario_st), language="text")
        with c_sap2: 
            st.caption("✈️ UNITARIO Vuelo (429)")
            st.code(fmt_sap(unitario_vuelo), language="text")
        with c_sap3: 
            st.caption("🧪 TOTAL Mezcla")
            st.code(fmt_sap(costo_mezcla_total), language="text")
        with c_sap4:
            st.markdown(f"<div style='background-color:#0d1b2a; padding:10px; border-radius:5px; border:1px solid #d4af37; text-align:center;'><p style='margin:0; color:#d4af37; font-size:12px;'>💰 COSTO x HA (Final)</p><h4 style='margin:0; color:white;'>$ {fmt_sap(costo_por_ha)}</h4></div>", unsafe_allow_html=True)

        # --- 3. TOTALES INFORMATIVOS ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("📊 **Resumen de Validación para SAP**")
        c_tot1, c_tot2, c_tot3 = st.columns(3)
        
        c_tot1.metric("Subtotal ST (459)", f"$ {fmt_sap(subtotal_st_finca)}")
        c_tot2.metric("Subtotal Vuelo (429)", f"$ {fmt_sap(subtotal_vuelo_finca)}")
        c_tot3.subheader(f"🔥 TOTAL: $ {fmt_sap(gran_total)}")
        
        st.markdown("---")
        st.markdown("### 🛰️ Coordenadas de Lanzamiento Final")
        
        tipo_mision = "DRONE" if mision_solo_dron else "AVION"
        
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            pistas_disponibles = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"]
            pista_manual = st.selectbox("📍 Confirmar Pista de Operación:", pistas_disponibles, index=pistas_disponibles.index(pista_sel) if pista_sel in pistas_disponibles else 0)

        with c_p2:
            st.info(f"🚀 Misión: {tipo_mision} | 📋 Referencia: {vuelo_ref}")
            
        if st.button("💾 DETONAR FACTURA Y GUARDAR EN BÓVEDA", type="primary", use_container_width=True):
            with st.spinner("🚀 Inyectando datos con Precisión de Francotirador..."):
                try:
                    if "gcp_credentials" in st.secrets:
                        cred_dict = dict(st.secrets["gcp_credentials"])
                        gc = gspread.service_account_from_dict(cred_dict)
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                    
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023")
                    hoja_maestra = boveda.worksheet("TABLA 1")
                    hoja_memoria = boveda.worksheet("MEMORIA")

                    fecha_str = fecha_operacion.strftime("%d/%m/%Y")
                    dia_sem = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"][fecha_operacion.weekday()]
                    num_sem = fecha_operacion.isocalendar()[1]
                    os_virtual = f"VIRT-{finca_limpia[:3]}-{datetime.now().strftime('%H%M')}"
                    
                    bloque_f = ""; sector_f = ""; ha_bruta_f = ""
                    if not df_t2.empty:
                        match_f = df_t2[df_t2.iloc[:, 0].str.upper().str.strip() == finca_limpia.upper().strip()]
                        if not match_f.empty:
                            sector_f = match_f.iloc[0, 1]
                            ha_bruta_f = match_f.iloc[0, 2]
                            bloque_f = match_f.iloc[0, 3]

                    # --- 🧮 CÁLCULOS MATEMÁTICOS DIRECTOS Y AJUSTE DE AERONAVE ---
                    ha_f = float(ha_dosis_final)
                    
                    # 🎯 INTELIGENCIA PROPORCIONAL: El horómetro se distribuye según las hectáreas de la finca actual
                    if mision_solo_dron:
                        h_total_v = ha_f / 10
                    else:
                        h_total_v = (ha_f / total_ha_cobro_escuadron) * horometro_final_avion if total_ha_cobro_escuadron > 0 else 0.0
                        
                    vol_total_gln = ha_f * 6
                    rend_min = h_total_v * 60
                    
                    piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                    
                    if mision_solo_dron:
                        if "DATAROT" in tipo_mision.upper(): hk_f = "DR51"
                        elif "GENESYS" in tipo_mision.upper(): hk_f = "DR52"
                        elif "AVIL" in tipo_mision.upper(): hk_f = "DR53"
                        else: hk_f = "DRONE_GEN"
                    else:
                        hk_f = hk_sel if 'hk_sel' in locals() else "AVION_REG"

                    modelo_f = "DRONE" if mision_solo_dron else "AVION"

                    try:
                        fila_excel = len(st.session_state['df_azul_actual']) + 2 
                    except:
                        fila_excel = 6 
                    
                    # 💰 ZONA FINANCIERA CALIBRADA (NETO VS COMERCIAL - UNIVERSAL) 🎯🎯
                    tarifa_vuelo_neta_ha = float(costo_neto_vuelo_total / total_ha_cobro_escuadron) if total_ha_cobro_escuadron > 0 else 0.0
                    valor_dominical = float(recargo_final)
                    
                    # PAGO A TERCEROS (AD): (Tarifa Neta + Recargo) * Hectáreas de la Finca Individual
                    total_pago_avion_neto = (tarifa_vuelo_neta_ha + valor_dominical) * ha_f
                    
                    # 📦 EMPAQUETADO MAESTRO - DETONACIÓN TABLA AZUL (34 Espacios)
                    row_azul = [""] * 34
                    row_azul[0] = os_virtual                  # A: ORDEN DE SERVICIO VIRTUAL
                    row_azul[1] = bloque_f
                    row_azul[2] = finca_limpia
                    row_azul[3] = sector_f
                    row_azul[4] = ha_bruta_f
                    row_azul[5] = ha_f                        # F: HECTÁREAS REALES
                    row_azul[6] = coctel_ganador
                    row_azul[7] = fecha_str
                    row_azul[8] = dia_sem
                    row_azul[9] = num_sem
                    # 🎯 AHORA SÍ: ODÓMETRO CON 2 DECIMALES EXACTOS (Columna K)
                    row_azul[10] = round(h_total_v, 2)
                    row_azul[11] = 6
                    row_azul[12] = round(vol_total_gln, 2)    # M: VOLUMEN (gln)
                    row_azul[13] = round(h_total_v, 2)        # N: RENDIMIENTO (hora)
                    row_azul[14] = round(rend_min, 2)         # O: RENDIMIENTO (min)
                    row_azul[15] = piloto_f
                    row_azul[16] = hk_f
                    row_azul[17] = tipo_mision
                    
                    row_azul[18] = float(gran_total)               # S: COSTO AVIÓN ($) [Total general facturado]
                    row_azul[19] = round(tarifa_vuelo_neta_ha, 2)  # T: COSTO AVIÓN ($/ha) [NETO PURO 2 DEC] 
                    row_azul[20] = round(valor_dominical, 2)       # U: DOMINIC ($/ha) [NETO PURO 2 DEC]
                    row_azul[21] = float(gran_total)               # V: COSTO AVIÓN ($/finca)
                    row_azul[23] = pista_manual                    # X: PISTA
                    row_azul[28] = float(gran_total)               # AC: COSTO TOTAL
                    row_azul[29] = round(total_pago_avion_neto, 2) # AD: TOTAL PAGO AVIÓN [NETO EXACTO 2 DEC] 
                    row_azul[32] = tipo_productor                  # AG: TIPO DE PRODUCTOR
                    row_azul[33] = "GÉNESIS_V2_PRO"                # AH: SELLO DE SISTEMA
                    
                    # 📦 EMPAQUETADO APOYO2023
                    fila_apoyo = [""] * 15
                    fila_apoyo[0] = "=IFERROR(ROW()-3, 0)" 
                    fila_apoyo[1] = finca_limpia
                    fila_apoyo[2] = ha_f
                    fila_apoyo[3] = float(costo_por_ha)
                    fila_apoyo[5] = fecha_str
                    fila_apoyo[8] = coctel_ganador
                    fila_apoyo[10] = pista_manual
                    fila_apoyo[13] = tipo_mision
                    
                    # 🔥 ESTRATEGIA FRANCOTIRADOR V3 (Escáner Infrarrojo + Expansión Automática)
                    col_azul = hoja_maestra.col_values(1)
                    fila_destino_azul = 1
                    for i in range(len(col_azul)-1, -1, -1):
                        if str(col_azul[i]).strip() != "":
                            fila_destino_azul = i + 2
                            break
                    
                    if fila_destino_azul > hoja_maestra.row_count:
                        hoja_maestra.add_rows(10)

                    col_apoyo = hoja_apoyo.col_values(2)
                    fila_destino_apoyo = 1
                    for i in range(len(col_apoyo)-1, -1, -1):
                        if str(col_apoyo[i]).strip() != "":
                            fila_destino_apoyo = i + 2
                            break
                            
                    if fila_destino_apoyo > hoja_apoyo.row_count:
                        hoja_apoyo.add_rows(10)

                    fila_apoyo[0] = fila_destino_apoyo - 3

                    hoja_maestra.update(range_name=f"A{fila_destino_azul}", values=[row_azul], value_input_option='USER_ENTERED')
                    hoja_apoyo.update(range_name=f"A{fila_destino_apoyo}", values=[fila_apoyo], value_input_option='USER_ENTERED')
                    
                    # 🧪 DESEMBARCO DE QUÍMICOS EN LA PESTAÑA 'MEMORIA' (LÓGICA SAMURAI)
                    try:
                        # 1. Escáner Anti-Duplicados (Lectura Rápida de la Bóveda)
                        datos_memoria = hoja_memoria.get_all_values()
                        set_existentes = set()
                        if len(datos_memoria) > 1:
                            for r in datos_memoria[1:]:
                                if len(r) >= 10:
                                    # Llave de seguridad: FECHA | FINCA | PRODUCTO
                                    llave = f"{str(r[0]).strip()}|{str(r[9]).strip().upper()}|{str(r[3]).strip().upper()}"
                                    set_existentes.add(llave)
                        
                        # 2. Asignación Dinámica de Bodega (Filtro Avión/Dron)
                        bodega_f = "BODEGA PRINCIPAL DRON" if mision_solo_dron else "BODEGA PRINCIPAL AVIÓN"
                        
                        filas_memoria = []
                        contador_nuevos = 0
                        
                        # 3. Bucle de Inyección
                        for idx, row in edited_df.iterrows():
                            nombre_prod = str(row.get("A: Producto", "")).strip().upper()
                            if "⚠️" not in nombre_prod and nombre_prod != "" and nombre_prod != "NAN":
                                
                                # Comprobamos si ya existe en SAP/Memoria
                                llave_actual = f"{fecha_str}|{finca_limpia}|{nombre_prod}"
                                
                                if llave_actual not in set_existentes:
                                    dosis_prod = float(row.get("D: Dosis Total (Sistema)", 0))
                                    lote_prod = str(row.get("G: Lotes (SAP)", "S/N"))
                                    
                                    fila_m = [""] * 10
                                    fila_m[0] = fecha_str                                      # A: FECHA
                                    fila_m[1] = coctel_ganador                                 # B: ORDEN/COCTEL
                                    fila_m[2] = str(pista_manual).split("-")[0].strip()[:4]    # C: PISTA
                                    fila_m[3] = nombre_prod                                    # D: PRODUCTO
                                    fila_m[4] = lote_prod                                      # E: LOTE
                                    fila_m[5] = float(dosis_prod)                              # F: CANTIDAD
                                    fila_m[6] = bodega_f                                       # G: BODEGA (Dinámica)
                                    fila_m[7] = ""                                             # H: MODELO (Vacío)
                                    fila_m[8] = "X"                                            # I: Facturado
                                    fila_m[9] = finca_limpia                                   # J: FINCA
                                    
                                    filas_memoria.append(fila_m)
                                    contador_nuevos += 1
                                    
                        # 4. Detonación por Lotes (Protege la API de Google)
                        if filas_memoria:
                            hoja_memoria.append_rows(filas_memoria, value_input_option='USER_ENTERED')
                            st.toast(f"💾 Memoria Samurai: {contador_nuevos} productos nuevos guardados.")
                        else:
                            st.toast("⚠️ Memoria Samurai: Los productos ya existían. No se duplicaron.")
                            
                    except Exception as e_mem:
                        st.warning(f"⚠️ Nota de sistema: Error al guardar en MEMORIA: {e_mem}")

                    st.balloons()
                    st.success(f"✅ IMPACTO TOTAL CONFIRMADO. Referencia {os_virtual} inyectada exactamente en la fila {fila_destino_azul}.")
                    
                    if 'memoria_excel' in st.session_state:
                        del st.session_state['memoria_excel']

                except Exception as e_save:
                    st.error(f"🚨 Falla en el Gatillo de Guardado: {e_save}")
                    
# =====================================================================
# ⌨️ 4. INGRESO MANUAL ACELERADO Y LEGALIZACIÓN (OS)
# =====================================================================
elif menu == "⌨️ 4. Ingreso Manual Acelerado (OS)":
    st.markdown("<h1 class='titulo-principal'>Gestión y Legalización de Órdenes (OS)</h1>", unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📝 1. Ingreso OS Manual (Desde Cero)", "🔄 2. Legalizar Vuelos Virtuales (Automático)"])

    # -----------------------------------------------------------------
    # PESTAÑA 1: INGRESO MANUAL ACELERADO (V3)
    # -----------------------------------------------------------------
    with tab1:
        st.subheader("Puesto de Control y Digitación Rápida")
        col_ref1, col_ref2 = st.columns([3, 1])
        with col_ref2:
            if st.button("🔄 RECARGAR BASES", use_container_width=True, key="btn_recargar_m4"):
                st.session_state.pop('memoria_excel', None)
                st.rerun()

        try:
            if "gcp_credentials" in st.secrets:
                gc1 = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
            else:
                gc1 = gspread.service_account(filename='credenciales.json')
            
            boveda1 = gc1.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            hoja_maestra1 = boveda1.worksheet("TABLA 1")
            
            if 'memoria_excel' not in st.session_state:
                with st.spinner("📡 Sincronizando Cerebro (Pilotos, Aviones y Apoyo)..."):
                    memoria = {}
                    memoria['col_os'] = hoja_maestra1.col_values(1)
                    
                    pilotos_raw = hoja_maestra1.col_values(16)
                    memoria['lista_pilotos'] = sorted(list(set([str(p).strip().upper() for p in pilotos_raw if p and str(p).upper() not in ["PILOTO", "PILOTO AVIÓN"]])))
                    
                    ws_t2_1 = boveda1.worksheet("TABLA 2")
                    d_t2_1 = ws_t2_1.get_all_values()
                    d_t2_limpio = [r + [""] * (12 - len(r)) if len(r) < 12 else r for r in d_t2_1]
                    memoria['df_t2'] = pd.DataFrame(d_t2_limpio[4:]) 
                    memoria['lista_hks'] = sorted(list(set([str(r[8]).strip().upper() for r in d_t2_limpio[4:] if r[8]])))

                    ws_ap_1 = boveda1.worksheet("TABLA DE APOYO2023")
                    d_ap_1 = ws_ap_1.get_all_values()
                    memoria['df_apoyo'] = pd.DataFrame(d_ap_1)
                    
                    st.session_state['memoria_excel'] = memoria

            mem = st.session_state['memoria_excel']
            lista_os_existentes = [str(os).strip() for os in mem['col_os'] if str(os).strip() != ""]
            df_t2_m4 = mem['df_t2']
            df_apoyo_m4 = mem['df_apoyo']
            
            lista_fincas_oficiales = sorted(list(set([str(f).strip().upper() for f in df_t2_m4.iloc[:, 0] if f])))
            lista_cocteles_oficiales = sorted(list(set([str(c).strip() for c in hoja_maestra1.col_values(7) if c and c != "COCTEL"])))

        except Exception as e:
            st.error(f"🚨 Error de enlace: {e}")
            st.stop()

        st.markdown("---")
        with st.expander("📝 1. DATOS DE LA ORDEN", expanded=True):
            c1, c2, c3 = st.columns(3)
            os_val = c1.text_input("Nº Orden (Ej: 318)", key="os_manual")
            fecha_dt = c2.date_input("📅 Fecha de Operación", format="DD/MM/YYYY", key="fecha_manual")
            piloto_val = c3.selectbox("👨‍✈️ Piloto", ["---"] + mem.get('lista_pilotos', []), key="piloto_manual")
            
            c4, c5, c6 = st.columns(3)
            hk_val = c4.selectbox("✈️ Matrícula (HK)", ["---"] + mem.get('lista_hks', []), key="hk_manual")
            horo_val = st.text_input("⏱️ Horómetro TOTAL (Ej: 1.5)", value="0", key="horo_manual")
            costo_val = st.text_input("💵 Tarifa / Ha", value="0", key="costo_manual")
            recargo_val = st.text_input("➕ Recargo Unitario ($)", value="0", key="recargo_manual")

        st.markdown("### 📍 2. FINCAS Y HECTÁREAS")
        st.info("💡 Si deja el Cóctel en blanco, Génesis lo buscará por FECHA y FINCA en la Tabla de Apoyo.")
        
        df_fincas_vacio = pd.DataFrame([{"nombre_finca": "", "hectareas": 0.0, "coctel": ""}])
        df_editado = st.data_editor(
            df_fincas_vacio, use_container_width=True, num_rows="dynamic", key="editor_manual",
            column_config={
                "nombre_finca": st.column_config.SelectboxColumn("Finca Oficial", options=lista_fincas_oficiales, required=True),
                "coctel": st.column_config.SelectboxColumn("Cóctel (Opcional)", options=lista_cocteles_oficiales),
                "hectareas": st.column_config.NumberColumn("Ha", format="%.2f", required=True)
            }
        )

        if st.button("🚀 PROCESAR E INYECTAR DATOS", type="primary", use_container_width=True, key="btn_inyect_manual"):
            if not os_val or piloto_val == "---" or hk_val == "---":
                st.error("🚨 Faltan datos críticos.")
            elif str(os_val).strip() in lista_os_existentes:
                st.error("🚨 Esta OS ya fue inyectada anteriormente.")
            else:
                try:
                    with st.spinner("🧠 El Transportador está cruzando datos..."):
                        f_str = fecha_dt.strftime("%d/%m/%Y")
                        
                        mod_av = ""; pist_av = ""
                        match_av = df_t2_m4[df_t2_m4.iloc[:, 8].str.strip() == hk_val]
                        if not match_av.empty:
                            mod_av, pist_av = match_av.iloc[0, 9], match_av.iloc[0, 10]

                        filas_finales = []
                        t_ha_os = sum(df_editado['hectareas'])
                        
                        h_tot = float(str(horo_val).replace(',','.'))
                        p_tar = float(str(costo_val).replace(',','.'))
                        p_rec = float(str(recargo_val).replace(',','.'))

                        for _, f in df_editado.iterrows():
                            n_finca = str(f['nombre_finca']).upper().strip()
                            if not n_finca: continue
                            
                            bloq = ""; sect = ""; hab = 0; t_prod = ""
                            m_f = df_t2_m4[df_t2_m4.iloc[:, 0].str.upper().str.strip() == n_finca]
                            if not m_f.empty:
                                sect, hab, bloq, t_prod = m_f.iloc[0, 1], extraer_numero(m_f.iloc[0, 2]), m_f.iloc[0, 3], m_f.iloc[0, 5]
                            
                            coctel_final = str(f.get('coctel', '')).strip()
                            if not coctel_final or coctel_final == "None" or coctel_final == "":
                                mask = (df_apoyo_m4.iloc[:, 1].str.upper().str.strip() == n_finca) & \
                                       (df_apoyo_m4.iloc[:, 5].str.strip() == f_str)
                                match_ap = df_apoyo_m4[mask]
                                
                                if not match_ap.empty:
                                    coctel_final = match_ap.iloc[0, 8]
                                else:
                                    match_hist = df_apoyo_m4[df_apoyo_m4.iloc[:, 1].str.upper().str.strip() == n_finca]
                                    if not match_hist.empty: coctel_final = match_hist.iloc[-1, 8]

                            ha_n = float(f['hectareas'])
                            h_prop = (ha_n / t_ha_os) * h_tot if t_ha_os > 0 else 0
                            costo_f = (ha_n * p_tar) + (ha_n * p_rec)
                            
                            row = [""] * 34
                            row[0], row[1], row[2], row[3], row[4], row[5] = os_val, bloq, n_finca, sect, hab, ha_n
                            row[6], row[7], row[8], row[9] = coctel_final, f_str, fecha_dt.strftime("%A"), fecha_dt.isocalendar()[1]
                            row[10], row[11], row[13], row[15], row[16] = h_tot, 6, round(h_prop,2), piloto_val, hk_val
                            row[17], row[18], row[19], row[20], row[21], row[23] = mod_av, round(costo_f,2), p_tar, p_rec, round(costo_f,2), pist_av
                            row[28], row[32], row[33] = round(ha_n * p_tar,2), t_prod, "GENESIS_INTELIGENTE"
                            
                            # 🔥 PYTHON TOMA EL MANDO: Inyección de Fórmulas Inteligentes
                            row[24] = '=INDIRECT("Y"&(ROW()-1))'  # Arrastra el incremento
                            row[25] = '=INDIRECT("Z"&(ROW()-1))'  # Arrastra el límite
                            row[26] = '=IFERROR(INDIRECT("S"&ROW())/INDIRECT("F"&ROW()), 0)' # Calcula Costo/Ha
                            row[27] = '=IF(INDIRECT("AA"&ROW())>INDIRECT("Z"&ROW()), "SUPERIOR", "INFERIOR")' # Evalúa Alerta
                            row[30] = '=INDIRECT("AE"&(ROW()-1))' # Arrastra Col AE
                            
                            filas_finales.append(row)
                        
                        if filas_finales:
                            hoja_maestra1.append_rows(filas_finales, value_input_option='USER_ENTERED')
                            st.balloons()
                            st.success(f"🎯 ¡OPERACIÓN EXITOSA! OS {os_val} inyectada con Cóctel y Fórmulas Automáticas.")
                            st.session_state.pop('memoria_excel', None) 
                        
                except Exception as e: st.error(f"Error en inyección: {e}")

    # -----------------------------------------------------------------
    # PESTAÑA 2: ESCÁNER DE LEGALIZACIÓN MULTI-OS
    # -----------------------------------------------------------------
    with tab2:
        st.markdown("### 🔄 Escáner de Legalización Multi-OS")
        
        if "gcp_credentials" in st.secrets:
            gc2 = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
        else:
            gc2 = gspread.service_account(filename='credenciales.json')
        
        sh2 = gc2.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        ws_t1_2 = sh2.worksheet("TABLA 1")
        ws_apoyo_2 = sh2.worksheet("TABLA DE APOYO2023")
        
        with st.spinner("Escaneando TABLA 1 en busca de misiones por legalizar..."):
            datos_t1 = ws_t1_2.get_all_values()
            pendientes = []
            for idx, row in enumerate(datos_t1[5:]):
                if len(row) > 19:
                    os_val = str(row[0]).upper()
                    equipo = str(row[17]).upper() 
                    if os_val.startswith("VIRT-") and ("AVION" in equipo or equipo == ""):
                        pendientes.append({
                            "fila_real": idx + 6,
                            "os_virt": os_val,
                            "finca": row[2],
                            "ha": extraer_numero(row[5]),
                            "costo_ha": extraer_numero(row[19]), 
                            "total": extraer_numero(row[18]),
                            "modelo": equipo
                        })

        if not pendientes:
            st.success("✅ No hay misiones de Avión pendientes por legalizar. ¡Cielo despejado!")
        else:
            df_pend = pd.DataFrame(pendientes)
            opciones_virt = df_pend.apply(lambda x: f"Fila {x['fila_real']} | {x['finca']} | {x['ha']} Ha | {x['os_virt']}", axis=1).tolist()
            seleccion = st.selectbox("🎯 Seleccione Vuelo Virtual para Legalizar:", opciones_virt)
            
            vuelo_sel = df_pend.iloc[opciones_virt.index(seleccion)]
            
            st.markdown("---")
            st.subheader(f"🛠️ Desglose de OS para: {vuelo_sel['finca']}")
            
            datos_apoyo = ws_apoyo_2.get_all_values()
            lista_todas_fincas = sorted(list(set([r[1] for r in datos_apoyo[3:] if len(r)>1 and r[1]])))

            if 'legalizador_rows' not in st.session_state:
                st.session_state.legalizador_rows = [{"OS_Real": "", "Finca": vuelo_sel['finca'], "Hectáreas": vuelo_sel['ha'], "Costo_Ha": vuelo_sel['costo_ha']}]

            col_btn, _ = st.columns([1, 4])
            if col_btn.button("➕ Añadir Finca/OS al Combo"):
                st.session_state.legalizador_rows.append({"OS_Real": "", "Finca": "", "Hectáreas": 0.0, "Costo_Ha": 0.0})
                st.rerun()

            rows_finales = []
            total_ha_asignadas = 0.0

            for i, row in enumerate(st.session_state.legalizador_rows):
                with st.container(border=True):
                    c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                    os_r = c1.text_input(f"OS Real #{i+1}", value=row["OS_Real"], key=f"os_r_{i}")
                    finca_r = c2.selectbox(f"Finca #{i+1}", [""] + lista_todas_fincas, 
                                           index=lista_todas_fincas.index(row["Finca"])+1 if row["Finca"] in lista_todas_fincas else 0, key=f"f_r_{i}")
                    
                    costo_sugerido = row["Costo_Ha"]
                    if finca_r != row["Finca"] and finca_r != "":
                        for r_ap in reversed(datos_apoyo):
                            if len(r_ap)>3 and r_ap[1] == finca_r:
                                costo_sugerido = extraer_numero(r_ap[3])
                                break
                    
                    ha_r = c3.number_input(f"Ha #{i+1}", value=float(row["Hectáreas"]), key=f"h_r_{i}")
                    costo_r = c4.number_input(f"$/Ha #{i+1}", value=float(costo_sugerido), key=f"c_r_{i}")
                    
                    rows_finales.append({"OS": os_r, "Finca": finca_r, "Ha": ha_r, "Costo": costo_r})
                    if finca_r == vuelo_sel['finca']: total_ha_asignadas += ha_r

            st.markdown("---")
            diferencia = round(vuelo_sel['ha'] - total_ha_asignadas, 2)
            
            c_m1, c_m2 = st.columns(2)
            c_m1.metric("🚜 Ha Objetivo (Finca Original)", f"{vuelo_sel['ha']} Ha")
            c_m2.metric("⚖️ Diferencia Pendiente", f"{diferencia} Ha", delta=-diferencia, delta_color="inverse")

            if st.button("🚀 DETONAR LEGALIZACIÓN EN TABLA 1", type="primary", use_container_width=True):
                if abs(diferencia) > 0.05:
                    st.error(f"❌ Error de cuadre: Aún faltan {diferencia} Ha por asignar.")
                else:
                    try:
                        with st.spinner("Legalizando y respetando Fórmulas MAP de Excel..."):
                            # 🔥 CORRECCIÓN CLAVE: Convertimos el número a "int" puro de Python
                            r_idx = int(vuelo_sel['fila_real'])
                            
                            Nuevas_Filas = []
                            for r_f in rows_finales:
                                fila_orig = datos_t1[r_idx - 1]
                                nueva = list(fila_orig) # Copiamos la fila original completa
                                
                                # Actualizamos asegurando que todos sean datos nativos
                                nueva[0] = str(r_f["OS"])       # A: OS
                                nueva[2] = str(r_f["Finca"])    # C: Finca
                                nueva[5] = float(r_f["Ha"])       # F: Ha
                                nueva[19] = float(r_f["Costo"])   # T: Costo Ha
                                nueva[18] = float(round(r_f["Ha"] * r_f["Costo"], 0)) # S: Total
                                nueva[21] = nueva[18]      # V: Subtotal
                                
                                # 🔥 REGLA DE ORO MAP: Dejamos completamente vacías las columnas de fórmulas
                                # Y=24, Z=25, AA=26, AB=27, AE=30
                                indices_vacios = [24, 25, 26, 27, 30]
                                for idx_v in indices_vacios:
                                    if idx_v < len(nueva): 
                                        nueva[idx_v] = ""
                                
                                Nuevas_Filas.append(nueva)

                            # Borramos la fila VIRT- original e insertamos las nuevas (limpias)
                            ws_t1_2.delete_rows(r_idx)
                            ws_t1_2.insert_rows(Nuevas_Filas, r_idx, value_input_option='USER_ENTERED')
                            
                            st.balloons()
                            st.success(f"🎯 LEGALIZACIÓN PERFECTA. Python se apartó y dejó que su fórmula MAP hiciera el cálculo.")
                            del st.session_state.legalizador_rows
                            st.rerun()
                    except Exception as e:
                        st.error(f"🚨 Falla en el sistema: {e}")
# =====================================================================
# 📈 5. SINCRONIZACIÓN PRECIOS Y TARIFARIO MAESTRO
# =====================================================================
elif menu == "📈 5. Sincronización Precios":
    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    
    # --- 🧮 NUEVA SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Copia Rápida)")
        st.info("💡 Obtenga la lista de precios exactos multiplicados por el margen de cada perfil, listos para copiar y pegar en SAP.")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Conectando con la Bóveda de Configuración..."):
                try:
                    if "gcp_credentials" in st.secrets:
                        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                        
                    sh_gen = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    raw_config = sh_gen.worksheet("Configuración").get_all_values()
                    
                    lista_precios = []
                    for row in raw_config:
                        if len(row) > 9:
                            prod = str(row[8]).upper().strip()
                            if prod and prod != "PRODUCTO" and "INVENTARIO" not in prod:
                                costo_base = extraer_numero(row[9])
                                if costo_base > 0:
                                    lista_precios.append({
                                        "PRODUCTO": prod,
                                        "COSTO BASE": costo_base,
                                        "TERCERO (+45.1%)": round(costo_base * 1.451, 0),
                                        "AFILIADO (+16.4%)": round(costo_base * 1.164, 0),
                                        "COOPERATIVA / SOCIO (+11.2%)": round(costo_base * 1.112, 0),
                                        "ORGÁNICO (+1.1%)": round(costo_base * 1.011, 0)
                                    })
                    
                    if lista_precios:
                        # 🎯 ORDENAMIENTO MILITAR: Alfabético por NOMBRE
                        df_tarifario = pd.DataFrame(lista_precios).sort_values(by="PRODUCTO").reset_index(drop=True)
                        st.session_state['df_tarifario'] = df_tarifario
                        st.success(f"✅ Tarifario cargado: {len(lista_precios)} productos ordenados alfabéticamente (A-Z).")
                    else:
                        st.warning("⚠️ El escáner no encontró productos con precios válidos.")
                except Exception as e:
                    st.error(f"🚨 Error al generar tarifario: {e}")
                    
        # 🛡️ SEGURO DE VIDA: Solo renderiza si el DataFrame existe y no está vacío
        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_t = st.session_state['df_tarifario']
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Copia Individual (Por Producto)"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes (Ordenada por Producto)")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### Caja de Copiado Masivo (Formación Alineada)")
                col_margen = st.selectbox("1️⃣ Seleccione el Perfil de Productor:", 
                                          ["TERCERO (+45.1%)", "AFILIADO (+16.4%)", "COOPERATIVA / SOCIO (+11.2%)", "ORGÁNICO (+1.1%)", "COSTO BASE"])
                
                incluir_nombres = st.toggle("🔘 Incluir Nombre del Producto (Alineación Perfecta)", value=False)
                st.caption(f"2️⃣ Copie la lista haciendo clic en el ícono de la esquina superior derecha:")
                
                if col_margen in df_t.columns:
                    if incluir_nombres:
                        # 🎯 JUSTIFICADOR MATEMÁTICO: Encuentra el nombre más largo para alinear
                        max_len = df_t["PRODUCTO"].apply(len).max() + 4
                        
                        lista_textos = []
                        for _, row in df_t.iterrows():
                            nombre = str(row["PRODUCTO"]).strip()
                            precio = fmt_sap(row[col_margen])
                            # Rellena con espacios a la derecha para emparejar la columna visualmente
                            nombre_alineado = nombre.ljust(max_len)
                            lista_textos.append(f"{nombre_alineado}\t{precio}")
                        texto_para_copiar = "\n".join(lista_textos)
                    else:
                        # Solo la columna de precios en Formato SAP (Ej: 76.041)
                        lista_textos = [fmt_sap(x) for x in df_t[col_margen]]
                        texto_para_copiar = "\n".join(lista_textos)
                        
                    st.code(texto_para_copiar, language="text")
                    
            with t3:
                st.markdown("#### Búsqueda y Copia Rápida Individual (Modo Francotirador)")
                prod_sel = st.selectbox("🔍 Buscar Producto Específico:", df_t["PRODUCTO"].tolist())
                
                if prod_sel:
                    datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
                    st.info(f"🎯 Valores calculados para: **{prod_sel}**")
                    
                    c1, c2, c3, c4, c5 = st.columns(5)
                    
                    with c1:
                        st.caption("Costo Base")
                        st.code(fmt_sap(datos_prod["COSTO BASE"]), language="text")
                    with c2:
                        st.caption("Orgánico")
                        st.code(fmt_sap(datos_prod["ORGÁNICO (+1.1%)"]), language="text")
                    with c3:
                        st.caption("Socio / Coop")
                        st.code(fmt_sap(datos_prod["COOPERATIVA / SOCIO (+11.2%)"]), language="text")
                    with c4:
                        st.caption("Afiliado")
                        st.code(fmt_sap(datos_prod["AFILIADO (+16.4%)"]), language="text")
                    with c5:
                        st.caption("Tercero")
                        st.code(fmt_sap(datos_prod["TERCERO (+45.1%)"]), language="text")
                        
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    semana_target = st.select_slider("Semana a actualizar:", options=list(range(1, 53)), value=19)

    if st.button("🚀 EJECUTAR OMEGA V12", use_container_width=True):
        try:
            with st.spinner(f"Sincronizando Semana {semana_target} al estilo Macro..."):
                if "gcp_credentials" in st.secrets:
                    cred_dict = dict(st.secrets["gcp_credentials"])
                    gc = gspread.service_account_from_dict(cred_dict)
                else:
                    gc = gspread.service_account(filename='credenciales.json')

                url_gen = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                sh_gen = gc.open_by_url(url_gen)
                
                raw_config = sh_gen.worksheet("Configuración").get_all_values(value_render_option='UNFORMATTED_VALUE')
                dict_precios = {}
                for row in raw_config:
                    if len(row) > 9:
                        prod = limpiar_texto_vba(row[8])
                        if prod and prod != "PRODUCTO":
                            dict_precios[prod] = val_seguro(row[9])

                raw_mezclas = sh_gen.worksheet("DD_Mesclas").get_all_values(value_render_option='UNFORMATTED_VALUE')
                dict_dosis = {}
                for row in raw_mezclas[12:]: 
                    if len(row) > 10:
                        prod_m = limpiar_texto_vba(row[9])
                        if prod_m:
                            dict_dosis[prod_m] = val_seguro(row[10])

                url_dest = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
                sh_dest = gc.open_by_url(url_dest)
                ws_datos = sh_dest.worksheet("DATOS")
                datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                
                col_semana = -1
                for i, v in enumerate(datos_dest[6]):
                    if str(v).strip() == str(semana_target):
                        col_semana = i + 1
                        break
                
                if col_semana == -1:
                    st.error(f"❌ No se halló la semana {semana_target} en la Fila 7.")
                else:
                    updates = []
                    for r_idx, row in enumerate(datos_dest):
                        n_fila = r_idx + 1
                        if n_fila < 8 or len(row) < 4: continue
                        
                        tipo_tabla = limpiar_texto_vba(row[1]) 
                        producto_dest = limpiar_texto_vba(row[3])
                        
                        if not producto_dest: continue
                        
                        if producto_dest in dict_precios:
                            precio_unitario = dict_precios[producto_dest]
                            if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                if producto_dest in dict_dosis:
                                    dosis_valor = dict_dosis[producto_dest]
                                    valor_final = precio_unitario * dosis_valor
                                else:
                                    valor_final = precio_unitario
                            else:
                                valor_final = precio_unitario
                                
                            updates.append({
                                'range': gspread.utils.rowcol_to_a1(n_fila, col_semana),
                                'values': [[valor_final]]
                            })

                    if updates:
                        ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                        st.success(f"🎯 IMPACTO PERFECTO. {len(updates)} precios inyectados con precisión absoluta.")
                        st.balloons()
                    else:
                        st.warning("⚠️ No se encontraron productos coincidentes.")

        except Exception as e:
            st.error(f"🚨 FALLA DEL SISTEMA: {e}")
            
# =====================================================================
# ✈️ 6. RASTREO DOMINICALES
# =====================================================================
elif menu == "✈️ 6. Rastreo Dominicales":
    st.markdown("<h1 class='titulo-principal'>Rastreo e Inyección de Recargos</h1>", unsafe_allow_html=True)
    
    url_ori = st.text_input(
        "🔗 Pegue URL de GÉNESIS_OMEGA_V2_ESTABLE:", 
        placeholder="Pegue aquí el link..."
    )

    if st.button("🚀 RASTREAR E INYECTAR FALTANTES", use_container_width=True):
        if not url_ori or "http" not in url_ori:
            st.error("❌ Pegue una URL válida.")
        else:
            try:
                if "gcp_credentials" in st.secrets:
                    cred_dict = dict(st.secrets["gcp_credentials"])
                    gc = gspread.service_account_from_dict(cred_dict)
                else:
                    gc = gspread.service_account(filename='credenciales.json')
                    
                with st.spinner("Modo Inyección Exacta Activado..."):
                    url_dest = "https://docs.google.com/spreadsheets/d/1FTiKlHo2UF8lWHk4SrFf9oxTUa2Q_n1l5IK9XFoqQaU/edit"
                    
                    sh_dest = gc.open_by_url(url_dest)
                    ws_dest = sh_dest.sheet1
                    datos_dest = ws_dest.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    
                    max_f = datetime(1900, 1, 1)
                    dict_local = {}
                    
                    for i, row in enumerate(datos_dest):
                        row_padded = row + [""] * (5 - len(row)) if len(row) < 5 else row
                        if i + 1 >= 5 and str(row_padded[1]).strip() != "":
                            f_obj = procesar_fecha_pesada(row_padded[3])
                            if f_obj:
                                if f_obj > max_f: max_f = f_obj
                                dict_local[f"{str(row_padded[1]).strip().upper()}|{f_obj.date()}"] = i + 1

                    st.info(f"📅 Radar Destino: Última fecha validada -> {max_f.strftime('%d/%m/%Y')}")

                    sh_ori = gc.open_by_url(url_ori)
                    ws_ori = next((s for s in sh_ori.worksheets() if "TABLA 1" in s.title.upper()), sh_ori.sheet1)
                    
                    st.write("---")
                    st.write(f"👁️ **RAYOS X ACTIVADOS:** Leyendo Archivo: `{sh_ori.title}` | Pestaña: `{ws_ori.title}`")
                    
                    datos_ori = ws_ori.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    dict_nuevos = {}
                    memoria_fecha = None 
                    recargos_encontrados = 0
                    recargos_ignorados = 0
                    
                    for i, row in enumerate(datos_ori):
                        n_fila = i + 1
                        if n_fila < 6: continue
                        
                        row_padded = row + [""] * (25 - len(row)) if len(row) < 25 else row
                        
                        f_leida = procesar_fecha_pesada(row_padded[7])
                        if f_leida: 
                            memoria_fecha = f_leida 
                        
                        surcharge = limpiar_val_dom(row_padded[20])
                        
                        if surcharge > 0:
                            recargos_encontrados += 1
                            f_operacion = f_leida if f_leida else memoria_fecha
                            
                            if f_operacion and f_operacion > max_f:
                                finca = str(row_padded[2]).strip().upper() if row_padded[2] else "SIN FINCA"
                                ha = limpiar_val_dom(row_padded[5])
                                pista = str(row_padded[23]).strip().upper() if row_padded[23] else ""
                                
                                key = f"{finca}|{f_operacion.date()}"
                                
                                if key in dict_nuevos:
                                    dict_nuevos[key]['ha'] += ha
                                    if not dict_nuevos[key]['pista'] and pista: dict_nuevos[key]['pista'] = pista
                                else:
                                    f_formato = f"{['lunes','martes','miércoles','jueves','viernes','sábado','domingo'][f_operacion.weekday()]}, {['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][f_operacion.month-1]} {f_operacion.day}, {f_operacion.year}"
                                    dict_nuevos[key] = {
                                        'finca': finca, 'ha': ha, 'fec': f_formato,
                                        'sur': surcharge, 'pista': pista, 'semana': f_operacion.isocalendar()[1]
                                    }
                            else:
                                recargos_ignorados += 1

                    st.write(f"📊 **MÉTRICAS:** {recargos_encontrados} Recargos totales | {recargos_ignorados} Ignorados por fecha antigua.")
                    st.write("---")

                    if dict_nuevos:
                        prox_fila = len(datos_dest) + 1 
                        filas_nuevas = [[v['finca'], v['ha'], v['fec'], v['sur'], v['pista'], v['semana']] for v in dict_nuevos.values()]
                        ws_dest.update(f'B{prox_fila}', filas_nuevas, value_input_option='USER_ENTERED')
                        st.success(f"🎯 ¡IMPACTO PERFECTO! {len(filas_nuevas)} registros inyectados empezando en la fila {prox_fila}.")
                        st.balloons()
                    else:
                        st.warning("⚠️ El escáner vio los recargos, pero ninguno era posterior a la fecha del radar.")

            except Exception as e:
                st.error(f"🚨 FALLA DE SISTEMA: {type(e).__name__} - {str(e)}")

# =====================================================================
# ⚖️ 7. ARQUEO DE INVENTARIOS
# =====================================================================
# =====================================================================
# ⚖️ 7. ARQUEO DE INVENTARIOS
# =====================================================================
elif menu == "⚖️ 7. Arqueo de Inventarios":
    m7.ejecutar(quitar_tildes, purificar_lote)
# =====================================================================
# 📊 8. REPORTE TÁCTICO DE HECTÁREAS FUMIGADAS
# =====================================================================
elif menu == "📊 8. Reporte Hectáreas (Pistas)":
    m8.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada, HAS_MATPLOTLIB)


# =====================================================================
# 📈 9. DASHBOARD TÁCTICO (FUSIÓN EXCEL + STREAMLIT)
# =====================================================================
elif menu == "📈 9. Dashboard Táctico":
    m9.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada)

# =====================================================================
# 📊 MÓDULO 10: CENTRO DE INTELIGENCIA ESTRATÉGICA BI
# =====================================================================
elif menu == "📊 10. Inteligencia de Costos (BI)":
    st.markdown("<h1 class='titulo-principal'>📊 Centro de Inteligencia Estratégica BI</h1>", unsafe_allow_html=True)
    st.markdown("### 🛰️ Panel de Auditoría y Comportamiento Histórico por Finca")
    st.info("🤖 **MOTOR IA BI:** Extrayendo memoria histórica y datos vivos...")

    # 1. Limpiador Base
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
        
    # 2. 🎯 ESTANDARIZADOR BLINDADO
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
            # 🛡️ BLINDAJE ANTI-ÁREA BRUTA: Obligamos al escáner a saltarse la palabra "BRUTA"
            elif not area_ok and ('FUMIG' in col_u or ('AREA' in col_u and 'BRUTA' not in col_u) or col_u == 'HAS'):
                renombres[col] = 'AREA_MAESTRA'
                area_ok = True
                
        df.rename(columns=renombres, inplace=True)
        return df
        
    # 3. TRADUCTOR FINANCIERO
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

    with st.spinner("📡 Sincronizando Bóveda Maestra y Archivo Histórico (Motor Turbo)..."):
        try:
            # CANAL A: Datos Vivos (Conectado a Caché)
            url_act = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
            datos_brutos_act = descargar_matriz_rapida(url_act, "TABLA 1")
            
            if len(datos_brutos_act) > 5:
                df_vivos = pd.DataFrame(datos_brutos_act[5:], columns=datos_brutos_act[4])
                df_vivos = estandarizar_base(limpiar_encabezados(df_vivos))
                df_vivos['ORIGEN_BI'] = 'ACTUAL'
            else: df_vivos = pd.DataFrame()

            # CANAL B: Datos Históricos (Conectado a Caché)
            url_hist = "https://docs.google.com/spreadsheets/d/16OZdiWwW7nLHyZBEnhiKlDTDttR7Tjhn37O9zm6wJOk/edit"
            datos_brutos_hist = descargar_matriz_rapida(url_hist, "Datos")
            
            if len(datos_brutos_hist) > 0:
                df_historico = pd.DataFrame(datos_brutos_hist[1:], columns=datos_brutos_hist[0])
                df_historico = estandarizar_base(limpiar_encabezados(df_historico))
                df_historico['ORIGEN_BI'] = 'HISTORICO'
            else: df_historico = pd.DataFrame()

            # FUSIÓN DEFINITIVA
            if not df_vivos.empty and not df_historico.empty:
                columnas_comunes = list(set(df_vivos.columns).intersection(set(df_historico.columns)))
                if 'ORIGEN_BI' in columnas_comunes: columnas_comunes.remove('ORIGEN_BI')
                
                if 'COSTO_MAESTRO' in columnas_comunes and 'FINCA_MAESTRA' in columnas_comunes:
                    df_vivos_trim = df_vivos[columnas_comunes + ['ORIGEN_BI']].copy()
                    df_historico_trim = df_historico[columnas_comunes + ['ORIGEN_BI']].copy()
                    
                    super_base_bi = pd.concat([df_historico_trim, df_vivos_trim], ignore_index=True)
                    super_base_bi['FINCA_MAESTRA'] = super_base_bi['FINCA_MAESTRA'].astype(str).str.strip().str.upper()

                    # --- FASE 3: MOTOR DE TIEMPO Y FILTROS ---
                    st.markdown("---")
                    st.markdown("### 🎛️ Centro de Mando: Parámetros de Análisis")
                    
                    if 'FECHA_MAESTRA' in super_base_bi.columns:
                        super_base_bi['FECHA_DT'] = super_base_bi['FECHA_MAESTRA'].apply(procesar_fecha_pesada)
                        super_base_bi = super_base_bi.dropna(subset=['FECHA_DT'])
                        
                        super_base_bi['AÑO'] = super_base_bi['FECHA_DT'].dt.year.astype(int)
                        super_base_bi['MES'] = super_base_bi['FECHA_DT'].dt.month.astype(int)
                        super_base_bi['TRIMESTRE'] = super_base_bi['FECHA_DT'].dt.quarter.astype(int)
                        
                        fincas_disp = ["TODAS"] + sorted(super_base_bi['FINCA_MAESTRA'].dropna().unique().tolist())
                        años_disp = sorted(super_base_bi['AÑO'].unique().tolist(), reverse=True)
                        
                        col_modelo = 'MODELO' if 'MODELO' in super_base_bi.columns else None
                        if col_modelo:
                            super_base_bi[col_modelo] = super_base_bi[col_modelo].astype(str).str.strip().str.upper()
                            modelos_disp = ["TODOS"] + sorted(super_base_bi[col_modelo].unique().tolist())
                        else:
                            modelos_disp = ["TODOS"]
                        
                        f1, f2 = st.columns(2)
                        finca_sel = f1.selectbox("📍 Objetivo Geográfico (Finca)", fincas_disp)
                        modelo_sel = f2.selectbox("🚁 Escuadrón (Modelo/Tipo)", modelos_disp)
                        
                        t1, t2, t3, t4 = st.columns(4)
                        idx_base = 1 if len(años_disp) > 1 else 0
                        año_base = t1.selectbox("📅 Año Base (Referencia)", años_disp, index=idx_base)
                        año_comp = t2.selectbox("📆 Año Actual (Evaluar)", años_disp, index=0)
                        
                        tipo_periodo = t3.selectbox("⏱️ Lupa Temporal", ["AÑO COMPLETO", "POR TRIMESTRE", "POR MES"])
                        meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
                        
                        if tipo_periodo == "POR TRIMESTRE":
                            periodo_sel = t4.selectbox("📊 Seleccione Trimestre", [1, 2, 3, 4], format_func=lambda x: f"Q{x}")
                            etiq_periodo = f"Q{periodo_sel}"
                        elif tipo_periodo == "POR MES":
                            periodo_sel = t4.selectbox("📅 Seleccione Mes", list(meses_dict.keys()), format_func=lambda x: meses_dict[x])
                            etiq_periodo = meses_dict[periodo_sel]
                        else:
                            t4.markdown("<br><span style='color:gray;'>Visión Anual Activada</span>", unsafe_allow_html=True)
                            periodo_sel = "TODOS"
                            etiq_periodo = "Total"

                        df_finca = super_base_bi.copy()
                        if finca_sel != "TODAS": df_finca = df_finca[df_finca['FINCA_MAESTRA'] == finca_sel]
                        if col_modelo and modelo_sel != "TODOS": df_finca = df_finca[df_finca[col_modelo] == modelo_sel]
                            
                        df_finca['COSTO_NUM'] = df_finca['COSTO_MAESTRO'].apply(convertir_pesos)

                        df_periodo_a = df_finca[df_finca['AÑO'] == año_base].copy()
                        df_periodo_b = df_finca[df_finca['AÑO'] == año_comp].copy()
                        
                        if tipo_periodo == "POR TRIMESTRE":
                            df_periodo_a = df_periodo_a[df_periodo_a['TRIMESTRE'] == periodo_sel]
                            df_periodo_b = df_periodo_b[df_periodo_b['TRIMESTRE'] == periodo_sel]
                        elif tipo_periodo == "POR MES":
                            df_periodo_a = df_periodo_a[df_periodo_a['MES'] == periodo_sel]
                            df_periodo_b = df_periodo_b[df_periodo_b['MES'] == periodo_sel]

                        costo_a = df_periodo_a['COSTO_NUM'].mean() if not df_periodo_a.empty else 0
                        costo_b = df_periodo_b['COSTO_NUM'].mean() if not df_periodo_b.empty else 0
                        delta_pct = ((costo_b - costo_a) / costo_a * 100) if costo_a > 0 else 0
                        
                        # 7. Artillería Visual: Tarjetas de Impacto
                        st.markdown("### 📊 Auditoría de Costos: Impacto General por Hectárea")
                        
                        k1, k2, k3 = st.columns(3)
                        k1.metric(label=f"Costo Promedio Ha ({año_base})", value=f"$ {costo_a:,.0f}")
                        k2.metric(label=f"Costo Promedio Ha ({año_comp})", value=f"$ {costo_b:,.0f}")
                        k3.metric(label="Variación Total (%)", value=f"{delta_pct:+.2f} %", delta=f"{delta_pct:+.2f}%", delta_color="inverse")
                        
                        # --- 🚜 RESCATE DE HECTÁREAS (FILTRO ANTI-CLONES) ---
                        st.markdown("#### 🚜 Volumen Operativo (Hectáreas Aplicadas)")
                        col_area = 'AREA_MAESTRA' if 'AREA_MAESTRA' in df_finca.columns else None
                        
                        def limpiar_area(val):
                            try:
                                v = str(val).upper().replace(',', '.')
                                v = "".join([c for c in v if c.isdigit() or c == '.'])
                                return float(v) if v != '' else 0.0
                            except: return 0.0
                            
                        if col_area:
                            df_periodo_a.loc[:, 'AREA_NUM'] = df_periodo_a[col_area].apply(limpiar_area)
                            df_periodo_b.loc[:, 'AREA_NUM'] = df_periodo_b[col_area].apply(limpiar_area)
                            
                            area_a = df_periodo_a.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM'])['AREA_NUM'].sum() if not df_periodo_a.empty else 0.0
                            area_b = df_periodo_b.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM'])['AREA_NUM'].sum() if not df_periodo_b.empty else 0.0
                        else:
                            area_a, area_b = 0.0, 0.0

                        var_area = ((area_b - area_a) / area_a * 100) if area_a > 0 else 0

                        h1, h2, h3 = st.columns(3)
                        h1.metric(f"Total Hectáreas ({año_base})", f"{area_a:,.1f} Ha")
                        h2.metric(f"Total Hectáreas ({año_comp})", f"{area_b:,.1f} Ha")
                        
                        if area_a > 0:
                            h3.metric("Variación de Área", f"{var_area:+.1f} %", delta=f"{var_area:+.1f}%", delta_color="normal")
                        else:
                            h3.metric("Variación de Área", "N/A")
                        
                        # 8. Sistema de Alerta Temprana
                        st.markdown("<br>", unsafe_allow_html=True)
                        if delta_pct > 10:
                            st.error(f"⚠️ **ALERTA ROJA:** El costo operativo en {finca_sel} presenta una desviación del **{delta_pct:.1f}%**. Se requiere análisis de causa raíz.")
                        elif delta_pct < 0:
                            st.success(f"✅ **RENDIMIENTO ÓPTIMO:** El costo operativo se redujo. Excelente gestión logística.")
                        else:
                            st.info(f"⚖️ **ESTABILIDAD:** Los costos se mantienen dentro de los márgenes normales de variación.")
                            
                        # --- ⏱️ FRECUENCIA OPERATIVA ---
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("#### ⏱️ Análisis de Frecuencia: Ciclos Reales e Intervalo")
                        
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
                                diffs = [(inicios_ciclo[j] - inicios_ciclo[j-1]).days for j in range(1, ciclos)]
                                avg_int = sum(diffs) / len(diffs)
                            else:
                                avg_int = 0
                            return ciclos, avg_int
                            
                        ciclos_a, int_a = calcular_frecuencia(df_periodo_a)
                        ciclos_b, int_b = calcular_frecuencia(df_periodo_b)
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric(f"Ciclos ({año_base})", f"{ciclos_a} ciclos")
                        c2.metric(f"Ciclos ({año_comp})", f"{ciclos_b} ciclos", delta=f"{ciclos_b - ciclos_a} ciclos", delta_color="inverse")
                        
                        str_int_a = f"{int_a:.1f} días" if int_a > 0 else "N/A"
                        str_int_b = f"{int_b:.1f} días" if int_b > 0 else "N/A"
                        c3.metric(f"Intervalo Prom. ({año_base})", str_int_a)
                        
                        if int_a > 0 and int_b > 0:
                            delta_int = int_b - int_a
                            c4.metric(f"Intervalo Prom. ({año_comp})", str_int_b, delta=f"{delta_int:+.1f} días", delta_color="normal")
                        else:
                            c4.metric(f"Intervalo Prom. ({año_comp})", str_int_b)
                        
                        # --- 📊 FASE 4: VISORES GRÁFICOS Y ATRIBUCIÓN DE COSTOS ---
                        st.markdown("---")
                        st.markdown("### 🧬 Análisis de Causa Raíz: Atribución de Variaciones")
                        
                        df_tendencia = pd.concat([df_periodo_a, df_periodo_b])
                        if not df_tendencia.empty:
                            if tipo_periodo in ["AÑO COMPLETO", "POR TRIMESTRE"]:
                                tendencia_agrupa = df_tendencia.groupby(['AÑO', 'MES'])['COSTO_NUM'].mean().reset_index()
                                tendencia_agrupa['EJE_X'] = tendencia_agrupa['MES'].map(meses_dict)
                                tendencia_agrupa = tendencia_agrupa.sort_values('MES')
                                titulo_x = "Meses Operativos"
                            else:
                                df_tendencia['DIA'] = df_tendencia['FECHA_DT'].dt.day
                                tendencia_agrupa = df_tendencia.groupby(['AÑO', 'DIA'])['COSTO_NUM'].mean().reset_index()
                                tendencia_agrupa['EJE_X'] = "Día " + tendencia_agrupa['DIA'].astype(str)
                                tendencia_agrupa = tendencia_agrupa.sort_values('DIA')
                                titulo_x = f"Días Operativos ({etiq_periodo})"
                                
                            tendencia_agrupa['AÑO'] = tendencia_agrupa['AÑO'].astype(str)
                            fig_tendencia = px.line(
                                tendencia_agrupa, x='EJE_X', y='COSTO_NUM', color='AÑO', 
                                markers=True, color_discrete_sequence=['#2F75B5', '#ef4444']
                            )
                            fig_tendencia.update_layout(
                                yaxis_title="Costo Promedio ($ COP / Ha)", 
                                xaxis_title=titulo_x, 
                                plot_bgcolor='rgba(0,0,0,0)',
                                hovermode="x unified"
                            )
                            max_y = tendencia_agrupa['COSTO_NUM'].max() * 1.2
                            if not pd.isna(max_y): fig_tendencia.update_yaxes(range=[0, max_y])
                            fig_tendencia.update_traces(
                                line=dict(width=3), marker=dict(size=8),
                                texttemplate="$ %{y:,.0f}", textposition="top center",
                                hovertemplate="<b>%{x}</b><br>Costo: $ %{y:,.0f} COP/Ha<extra></extra>"
                            )
                            st.plotly_chart(fig_tendencia, use_container_width=True)
                        else:
                            st.warning("⚠️ No hay suficientes operaciones en este periodo exacto para trazar una curva comparativa.")
                            
                        st.markdown("<hr>", unsafe_allow_html=True)
                        
                        col_avion_ha = None
                        for col in df_finca.columns:
                            col_u = str(col).upper().replace('Ó', 'O')
                            if 'AVION' in col_u and ('/HA' in col_u or ' HA' in col_u or '(HA)' in col_u):
                                col_avion_ha = col
                                break
                        
                        if col_avion_ha:
                            df_periodo_a.loc[:, 'AVION_NUM'] = df_periodo_a[col_avion_ha].apply(convertir_pesos)
                            df_periodo_b.loc[:, 'AVION_NUM'] = df_periodo_b[col_avion_ha].apply(convertir_pesos)
                        else:
                            df_periodo_a.loc[:, 'AVION_NUM'] = 0.0
                            df_periodo_b.loc[:, 'AVION_NUM'] = 0.0

                        vuelo_a = df_periodo_a['AVION_NUM'].mean() if not df_periodo_a.empty else 0
                        vuelo_b = df_periodo_b['AVION_NUM'].mean() if not df_periodo_b.empty else 0
                        insumos_a = max(0, costo_a - vuelo_a)
                        insumos_b = max(0, costo_b - vuelo_b)

                        vuelo_tot_a = vuelo_a * area_a
                        vuelo_tot_b = vuelo_b * area_b
                        insumos_tot_a = insumos_a * area_a
                        insumos_tot_b = insumos_b * area_b

                        st.markdown("#### 🛩️ vs 🧪 Distribución del Encarecimiento")
                        categorias = [f'Análisis {año_base}', f'Análisis {año_comp}']
                        tab_unit, tab_glob = st.tabs(["🎯 Impacto Unitario (Promedio / Ha)", "💰 Impacto Global (Presupuesto Total)"])
                        
                        with tab_unit:
                            fig_unit = go.Figure(data=[
                                go.Bar(name='Costo Avión / Ha', x=categorias, y=[vuelo_a, vuelo_b], marker_color='#2F75B5', text=[f"$ {vuelo_a:,.0f}", f"$ {vuelo_b:,.0f}"], textposition='auto'),
                                go.Bar(name='Costo Insumos / Ha', x=categorias, y=[insumos_a, insumos_b], marker_color='#548235', text=[f"$ {insumos_a:,.0f}", f"$ {insumos_b:,.0f}"], textposition='auto')
                            ])
                            fig_unit.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor COP / Ha", margin=dict(t=20, b=20))
                            st.plotly_chart(fig_unit, use_container_width=True)
                            
                        with tab_glob:
                            fig_glob = go.Figure(data=[
                                go.Bar(name='Total Facturación Avión', x=categorias, y=[vuelo_tot_a, vuelo_tot_b], marker_color='#2F75B5', text=[f"$ {vuelo_tot_a:,.0f}", f"$ {vuelo_tot_b:,.0f}"], textposition='auto'),
                                go.Bar(name='Total Consumo Insumos', x=categorias, y=[insumos_tot_a, insumos_tot_b], marker_color='#548235', text=[f"$ {insumos_tot_a:,.0f}", f"$ {insumos_tot_b:,.0f}"], textposition='auto')
                            ])
                            fig_glob.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor Total COP", margin=dict(t=20, b=20))
                            st.plotly_chart(fig_glob, use_container_width=True)
                        
                        # 4. TABLA INTERACTIVA DE CÓCTELES
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("#### 📋 Desglose Operativo: Cócteles y Variación")
                        
                        col_coctel = 'COCTEL' if 'COCTEL' in df_finca.columns else ('COCTEL_MAESTRO' if 'COCTEL_MAESTRO' in df_finca.columns else None)
                        col_gln = 'GLN_HA' if 'GLN_HA' in df_finca.columns else None
                        
                        if col_coctel:
                            df_periodo_a.loc[:, col_coctel] = df_periodo_a[col_coctel].astype(str).str.strip().str.upper()
                            df_periodo_b.loc[:, col_coctel] = df_periodo_b[col_coctel].astype(str).str.strip().str.upper()
                            
                            agg_dict = {'COSTO_NUM': 'mean'}
                            if col_gln: agg_dict[col_gln] = 'mean'
                            
                            g_a = df_periodo_a.groupby(col_coctel).agg(agg_dict).reset_index()
                            g_b = df_periodo_b.groupby(col_coctel).agg(agg_dict).reset_index()
                            
                            tabla_autopsia = pd.merge(g_a, g_b, on=col_coctel, how='outer', suffixes=('_BASE', '_ACTUAL'))
                            tabla_autopsia.fillna(0, inplace=True)
                            
                            tabla_autopsia.rename(columns={
                                col_coctel: 'CÓCTEL APLICADO',
                                'COSTO_NUM_BASE': f'Costo/Ha ({año_base})',
                                'COSTO_NUM_ACTUAL': f'Costo/Ha ({año_comp})'
                            }, inplace=True)
                            
                            tabla_autopsia['Variación ($)'] = tabla_autopsia[f'Costo/Ha ({año_comp})'] - tabla_autopsia[f'Costo/Ha ({año_base})']
                            
                            if col_gln:
                                tabla_autopsia.rename(columns={
                                    f'{col_gln}_BASE': f'Gln/Ha ({año_base})',
                                    f'{col_gln}_ACTUAL': f'Gln/Ha ({año_comp})'
                                }, inplace=True)
                                
                            df_vista = tabla_autopsia.copy()
                            df_vista[f'Costo/Ha ({año_base})'] = df_vista[f'Costo/Ha ({año_base})'].map("$ {:,.0f}".format)
                            df_vista[f'Costo/Ha ({año_comp})'] = df_vista[f'Costo/Ha ({año_comp})'].map("$ {:,.0f}".format)
                            df_vista['Variación ($)'] = df_vista['Variación ($)'].map("$ {:,.0f}".format)
                            
                            st.dataframe(df_vista, use_container_width=True)
                            
                        # =====================================================================
                        # --- 🔬 NIVEL 2: ALGORITMO CHEF HÍBRIDO Y DELIBERADOR IA ---
                        # =====================================================================
                        st.markdown("<hr>", unsafe_allow_html=True)
                        st.markdown("### 🔬 Nivel 2: Composición del Cóctel y Variación Real de Insumos")

                        if col_coctel:
                            cocteles_disponibles = sorted(list(set(df_periodo_a[col_coctel].dropna().unique()) | set(df_periodo_b[col_coctel].dropna().unique())))
                            coctel_sel = st.selectbox("🎯 Seleccione un Cóctel para auditar su receta año vs año:", ["SELECCIONE UN CÓCTEL..."] + cocteles_disponibles)

                            if coctel_sel != "SELECCIONE UN CÓCTEL...":
                                with st.spinner("Desplegando Deliberador IA y conectando al Histórico de Precios..."):
                                    try:
                                        df_mezclas = pd.DataFrame()
                                        
                                        boveda_recetas = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                                        hoja_mezclas = boveda_recetas.worksheet("DD_Mesclas")
                                        data_mez = hoja_mezclas.get('A:D')
                                        if data_mez and len(data_mez) > 1:
                                            df_mezclas = pd.DataFrame(data_mez[1:], columns=data_mez[0])
                                            # Búsqueda a prueba de balas (sin espacios)
                                            df_mezclas['COCTEL_CLEAN'] = df_mezclas.iloc[:,0].astype(str).str.upper().str.replace(" ", "")
                                        
                                        data_conf = boveda_recetas.worksheet("Configuración").get_all_values()
                                        df_conf = pd.DataFrame(data_conf[1:], columns=data_conf[0])

                                        data_dicc = boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()
                                        df_dicc = pd.DataFrame(data_dicc[1:], columns=data_dicc[0])

                                        url_precios = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
                                        sh_precios = gc.open_by_url(url_precios)
                                        
                                        precios_consolidados = []
                                        for ws in sh_precios.worksheets():
                                            datos_hoja = ws.get_all_values()
                                            if not datos_hoja: continue
                                            
                                            idx_header, col_anio, col_prod = -1, -1, -1
                                            for i in range(min(10, len(datos_hoja))):
                                                fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                                                if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                                                    idx_header = i; col_anio = fila_upper.index('AÑO'); col_prod = fila_upper.index('PRODUCTO')
                                                    break
                                                    
                                            if idx_header != -1:
                                                for row in datos_hoja[idx_header+1:]:
                                                    if len(row) > max(col_anio, col_prod):
                                                        anio_str = str(row[col_anio]).strip()
                                                        prod_str = str(row[col_prod]).strip().upper()
                                                        if anio_str and prod_str:
                                                            col_inicio_semanas = max(col_anio, col_prod) + 1
                                                            valores_semana = [extraer_numero(v) for v in row[col_inicio_semanas:] if extraer_numero(v) > 0]
                                                            promedio = sum(valores_semana)/len(valores_semana) if valores_semana else 0.0
                                                            precios_consolidados.append({'AÑO': anio_str, 'PRODUCTO': prod_str, 'PRECIO_PROM': promedio})

                                        df_precios = pd.DataFrame(precios_consolidados)

                                        # =========================================================
                                        # =========================================================
                                        # 3. ALGORITMO CHEF HÍBRIDO: SOBERANÍA DE DD_Mesclas
                                        # =========================================================
                                        import re
                                        coctel_crudo = coctel_sel.upper().replace(" ", "")
                                        
                                        partes_coctel = coctel_crudo.split('+')
                                        base_coctel = partes_coctel[0]
                                        aditivos = partes_coctel[1:] if len(partes_coctel) > 1 else []

                                        match_num = re.search(r'\d+', base_coctel)
                                        dosis_aceite = int(match_num.group()) if match_num else 0
                                        solo_letras = re.sub(r'\d+', '', base_coctel)

                                        dict_prods_unicos = {}
                                        es_organico = False

                                        # SENSOR 1: PRODUCTOR ORGÁNICO (TABLA 2)
                                        try:
                                            data_t2 = boveda_recetas.worksheet("TABLA 2").get_all_values()
                                            df_t2 = pd.DataFrame(data_t2[1:], columns=data_t2[0])
                                            match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_sel.upper().strip()]
                                            if not match_f.empty and "ORGANIC" in str(match_f.iloc[0, 5]).upper():
                                                es_organico = True
                                        except: pass

                                        # 🧠 PRIORIDAD 1: CLONAR RECETA DESDE LA BÓVEDA (DD_Mesclas)
                                        receta_base = pd.DataFrame()
                                        if not df_mezclas.empty:
                                            # Intentamos buscar KM3 u Organizados
                                            if es_organico and not base_coctel.endswith('O'):
                                                coctel_prueba = f"{base_coctel}O"
                                                if not df_mezclas[df_mezclas['COCTEL_CLEAN'] == coctel_prueba].empty:
                                                    base_coctel = coctel_prueba
                                            
                                            receta_base = df_mezclas[df_mezclas['COCTEL_CLEAN'] == base_coctel]
                                            if receta_base.empty:
                                                receta_base = df_mezclas[df_mezclas['COCTEL_CLEAN'] == solo_letras]

                                        if not receta_base.empty:
                                            # Absorción total y fiel de sus datos de campo (Mantiene dosis reales de Acondicionador por pH/Dureza)
                                            for idx, row in receta_base.iterrows():
                                                prod = str(row.iloc[1]).strip().upper()
                                                dosis = extraer_numero(row.iloc[2])
                                                if dosis > 0 and prod not in ['NAN', '']:
                                                    dict_prods_unicos[prod] = dosis
                                        else:
                                            # 🧠 PRIORIDAD 2: RESPALDO DESDE EL DICCIONARIO (Solo si es un cóctel 100% nuevo)
                                            if not df_dicc.empty:
                                                siglas_validas = df_dicc[df_dicc['SIGLA'].astype(str).str.strip() != '']['SIGLA'].astype(str).str.strip().str.upper().unique().tolist()
                                                siglas_validas.sort(key=len, reverse=True)
                                                resto_letras = solo_letras
                                                for sigla in siglas_validas:
                                                    if sigla in resto_letras:
                                                        match_sig = df_dicc[df_dicc['SIGLA'].astype(str).str.strip().str.upper() == sigla]
                                                        if not match_sig.empty:
                                                            prod_name = str(match_sig.iloc[0]['PRODUCTO']).strip().upper()
                                                            dict_prods_unicos[prod_name] = extraer_numero(match_sig.iloc[0]['DOSIS'])
                                                        resto_letras = resto_letras.replace(sigla, '', 1)
                                            
                                            if dosis_aceite > 0: dict_prods_unicos['ACEITE DICAM'] = float(dosis_aceite)
                                            dict_prods_unicos['ACONDICIONADOR SV'] = 0.02 # Default de fábrica temporal
                                            dict_prods_unicos['ADHERENTE SV'] = 0.13

                                        # 🧠 FASE C: COMPLEMENTOS DE ADITIVOS (+ZN, +BT)
                                        if not df_dicc.empty:
                                            for ad in aditivos:
                                                match_sig = df_dicc[df_dicc['SIGLA'].astype(str).str.strip().str.upper() == ad]
                                                if not match_sig.empty:
                                                    prod_name = str(match_sig.iloc[0]['PRODUCTO']).strip().upper()
                                                    dict_prods_unicos[prod_name] = extraer_numero(match_sig.iloc[0]['DOSIS'])
                                                else:
                                                    if "ZN" in ad: dict_prods_unicos["ZINTRAC"] = 0.5
                                                    elif "BT" in ad: dict_prods_unicos["BANATREL"] = 0.5

                                        # 🧠 FASE D: COMPROBACIÓN DE LÍNEA DE CULTIVO DESDE DICCIONARIO
                                        if not df_dicc.empty:
                                            for prod_name in list(dict_prods_unicos.keys()):
                                                match_dicc = df_dicc[df_dicc['PRODUCTO'].astype(str).str.strip().str.upper() == prod_name]
                                                if not match_dicc.empty:
                                                    if 'ORGANIC' in str(match_dicc.iloc[0].get('TIPO DE CULTIVO', '')).upper(): es_organico = True

                                        # 🧠 FASE E: AUDITORÍA FINANCIERA-AGRONÓMICA (REGLAS DE ORO CORREGIDAS)
                                        
                                        # 1. El Aceite obedece al número de la sigla, pero conserva la marca de su hoja de mezclas
                                        if dosis_aceite > 0:
                                            aceite_key = next((k for k in dict_prods_unicos.keys() if "ACEITE" in k), "ACEITE DICAM")
                                            dict_prods_unicos[aceite_key] = float(dosis_aceite)
                                        else:
                                            # Si la sigla no tiene número de aceite, nos aseguramos de limpiarlo
                                            claves_aceite = [k for k in dict_prods_unicos.keys() if "ACEITE" in k]
                                            for k in claves_aceite: dict_prods_unicos.pop(k, None)

                                        # 2. Bifurcación Inteligente de Pegantes Obligatorios (Orgánico vs Convencional)
                                        if es_organico:
                                            # Orgánico: Remueve el adherente convencional y asegura el Sprayfix a 0.2
                                            claves_adherente = [k for k in dict_prods_unicos.keys() if "ADHERENTE" in k]
                                            for k in claves_adherente: dict_prods_unicos.pop(k, None)
                                            
                                            sprayfix_key = next((k for k in dict_prods_unicos.keys() if "SPRAYFIX" in k), "SPRAYFIX")
                                            if sprayfix_key not in dict_prods_unicos: dict_prods_unicos[sprayfix_key] = 0.2
                                        else:
                                            # Convencional: Remueve el Sprayfix orgánico y asegura el Adherente a 0.13
                                            claves_sprayfix = [k for k in dict_prods_unicos.keys() if "SPRAYFIX" in k]
                                            for k in claves_sprayfix: dict_prods_unicos.pop(k, None)
                                            
                                            adherente_key = next((k for k in dict_prods_unicos.keys() if "ADHERENTE" in k), "ADHERENTE SV")
                                            if adherente_key not in dict_prods_unicos: dict_prods_unicos[adherente_key] = 0.13

                                        # NOTA: El ACONDICIONADOR no se toca aquí. Se respeta la dosis exacta de su Excel.

                                        # Generación limpia de la lista final omitiendo dosis en cero
                                        prods_receta = [{"PRODUCTO": k, "DOSIS": v} for k, v in dict_prods_unicos.items() if v > 0]
                                        if prods_receta:
                                            matriz_mol = []
                                            
                                            def obtener_precio_promedio(producto, anio_obj):
                                                if not df_precios.empty:
                                                    match_df = df_precios[(df_precios['AÑO'] == str(anio_obj)) & (df_precios['PRODUCTO'] == producto)]
                                                    if match_df.empty:
                                                        match_df = df_precios[(df_precios['AÑO'] == str(anio_obj)) & (df_precios['PRODUCTO'].str.contains(producto))]
                                                    if not match_df.empty and match_df['PRECIO_PROM'].mean() > 0:
                                                        return match_df['PRECIO_PROM'].mean()
                                                
                                                if str(anio_obj) == str(año_comp) or str(anio_obj) == str(datetime.now().year):
                                                    match_conf = df_conf[df_conf.iloc[:, 8].astype(str).str.upper().str.strip() == producto]
                                                    if match_conf.empty:
                                                        match_conf = df_conf[df_conf.iloc[:, 8].astype(str).str.upper().str.strip().str.contains(producto)]
                                                    if not match_conf.empty: return extraer_numero(match_conf.iloc[0, 9])
                                                return 0.0

                                            costo_total_a = 0.0
                                            costo_total_b = 0.0

                                            for item in prods_receta:
                                                prod = item["PRODUCTO"]
                                                dosis = item["DOSIS"]
                                                precio_a = obtener_precio_promedio(prod, año_base)
                                                precio_b = obtener_precio_promedio(prod, año_comp)
                                                
                                                costo_ha_a = dosis * precio_a
                                                costo_ha_b = dosis * precio_b
                                                
                                                costo_total_a += costo_ha_a
                                                costo_total_b += costo_ha_b
                                                
                                                matriz_mol.append({
                                                    "INSUMO QUÍMICO": prod, "DOSIS/HA": f"{dosis:.3f}",
                                                    f"P. Prom. ({año_base})": f"$ {precio_a:,.0f}", f"P. Prom. ({año_comp})": f"$ {precio_b:,.0f}",
                                                    f"Costo/Ha ({año_base})": costo_ha_a, f"Costo/Ha ({año_comp})": costo_ha_b,
                                                    "Variación ($)": costo_ha_b - costo_ha_a
                                                })

                                            df_vista_mol = pd.DataFrame(matriz_mol).sort_values('Variación ($)', ascending=False)
                                            df_vista_mol[f"Costo/Ha ({año_base})"] = df_vista_mol[f"Costo/Ha ({año_base})"].map("$ {:,.0f}".format)
                                            df_vista_mol[f"Costo/Ha ({año_comp})"] = df_vista_mol[f"Costo/Ha ({año_comp})"].map("$ {:,.0f}".format)
                                            df_vista_mol["Variación ($)"] = df_vista_mol["Variación ($)"].map("$ {:,.0f}".format)
                                            
                                            st.dataframe(df_vista_mol, use_container_width=True, hide_index=True)
                                            
                                            c1, c2, c3 = st.columns(3)
                                            c1.metric(f"Total Teórico ({año_base})", f"$ {costo_total_a:,.0f}")
                                            c2.metric(f"Total Teórico ({año_comp})", f"$ {costo_total_b:,.0f}")
                                            c3.metric("Variación Cóctel", f"$ {costo_total_b - costo_total_a:,.0f}", delta=f"$ {costo_total_b - costo_total_a:,.0f}", delta_color="inverse")
                                            
                                            # =========================================================
                                            # 🤖 DELIBERADOR IA: INGENIERÍA INVERSA DE FACTURACIÓN
                                            # =========================================================
                                            if 'AVION_NUM' in df_periodo_b.columns:
                                                df_coctel_b = df_periodo_b[df_periodo_b[col_coctel] == coctel_sel]
                                                costo_total_facturado_b = df_coctel_b['COSTO_NUM'].mean() if not df_coctel_b.empty else 0
                                                vuelo_facturado_b = df_coctel_b['AVION_NUM'].mean() if not df_coctel_b.empty else 0
                                                insumos_facturados_b = max(0, costo_total_facturado_b - vuelo_facturado_b)
                                                
                                                if costo_total_b > 0 and insumos_facturados_b > 0:
                                                    diff_b = insumos_facturados_b - costo_total_b
                                                    
                                                    st.markdown("---")
                                                    st.markdown("### 🤖 Deliberador IA: Auditoría de Facturación SAP vs Receta Teórica")
                                                    
                                                    if abs(diff_b) <= 2000: # Tolerancia por redondeos de Excel
                                                        st.success(f"✅ **AUDITORÍA PERFECTA:** El costo de químicos facturados en SAP ($ {insumos_facturados_b:,.0f}) coincide con la receta ($ {costo_total_b:,.0f}).")
                                                    else:
                                                        st.warning(f"⚠️ **DISCREPANCIA DETECTADA:** Los insumos facturados ($ {insumos_facturados_b:,.0f}) no cuadran con el teórico ($ {costo_total_b:,.0f}). Diferencia: **$ {diff_b:,.0f} / Ha**")
                                                        
                                                        st.markdown("#### 🔍 Conclusiones del Deliberador:")
                                                        if diff_b > 0:
                                                            st.write(f"- 📈 **Sobrecosto:** Se cobró más de lo que indica la sigla. Es probable que se haya aplicado **SPRAYFIX**, **ADHERENTE** extra o mayor dosis de **ACEITE**.")
                                                        else:
                                                            st.write(f"- 📉 **Ahorro/Faltante:** Se cobró menos. Si la finca es orgánica, se facturó correctamente (sin adherente), o hubo un error a favor en SAP.")
                                                            
                                                        if not df_precios.empty:
                                                            st.write("- **Posibles causantes de la diferencia:**")
                                                            candidatos_encontrados = False
                                                            for idx, p_row in df_precios[df_precios['AÑO'] == str(año_comp)].iterrows():
                                                                precio_p = p_row['PRECIO_PROM']
                                                                for d in [0.02, 0.06, 0.13, 0.2, 0.5, 1.0, 2.0]:
                                                                    costo_teorico = precio_p * d
                                                                    if costo_teorico > 0 and abs(costo_teorico - abs(diff_b)) <= (abs(diff_b) * 0.15 + 500):
                                                                        st.info(f"💡 ¿Se aplicó/omitió **{p_row['PRODUCTO']}** a dosis de **{d} L/Ha**? (Costo aprox: $ {costo_teorico:,.0f})")
                                                                        candidatos_encontrados = True
                                                                        break
                                                            if not candidatos_encontrados:
                                                                st.write("No se detectó un químico individual que coincida exacto. Revise si hay una mezcla de aditivos omitidos.")

                                        else:
                                            st.info("No se encontraron ingredientes válidos para esta receta.")
                                    except Exception as e:
                                        st.error(f"🚨 Error en el cruce de históricos: {e}")
                        else:
                            st.warning("⚠️ No se encontró la columna 'COCTEL' en la base fusionada.")
# =====================================================================
                        # =====================================================================
                        # --- 🤝 NUEVO: SIMULADOR DE NEGOCIACIÓN Y AUDITORÍA DE TARIFAS ---
                        # =====================================================================
                        st.markdown("<hr>", unsafe_allow_html=True)
                        st.markdown("### 🤝 Simulador de Negociación (Tarifas de Aerofumigación)")
                        st.info("💡 RADAR BLINDADO: Extracción estricta de Tarifas Unitarias (Avión + Dominical). Lógica: (Nueva Tarifa Redondeada - Tarifa Actual Redondeada) × Hectáreas.")

                        # Filtros del simulador
                        c_sim1, c_sim2, c_sim3 = st.columns(3)

                        col_pista_sim = next((c for c in super_base_bi.columns if "PISTA" in c or "ALMACEN" in c), None)
                        if col_pista_sim:
                            pistas_sim_disp = ["TODAS"] + sorted(super_base_bi[col_pista_sim].dropna().astype(str).str.upper().unique().tolist())
                        else:
                            pistas_sim_disp = ["TODAS"]

                        sim_anio = c_sim1.selectbox("📅 Año a Auditar:", años_disp, key="sim_anio_v6")
                        sim_mes = c_sim2.selectbox("📆 Mes a Auditar:", list(meses_dict.keys()), format_func=lambda x: meses_dict[x], index=4) 
                        sim_pista = c_sim3.selectbox("📍 Base / Pista:", pistas_sim_disp, key="sim_pista_v6")

                        st.markdown("<br>", unsafe_allow_html=True)
                        c_sim_m1, c_sim_m2, c_sim_m3 = st.columns(3)
                        
                        margen_actual = c_sim_m1.number_input("📉 Margen Actual en Factura (%)", value=8.0, step=0.5, key="marg_act_v6")
                        margen_nuevo = c_sim_m2.number_input("📈 Nuevo Margen a Simular (%)", value=11.0, step=0.5, key="marg_nue_v6")
                        
                        with c_sim_m3:
                            st.markdown("<br>", unsafe_allow_html=True)
                            btn_simular = st.button("🚀 EJECUTAR SIMULACIÓN", type="primary", use_container_width=True, key="btn_simular_v6")

                        if btn_simular:
                            with st.spinner("Procesando auditoría con las columnas unitarias correctas..."):
                                df_sim = super_base_bi.copy()

                                df_sim = df_sim[df_sim['AÑO'] == int(sim_anio)]
                                df_sim = df_sim[df_sim['MES'] == int(sim_mes)]
                                if col_pista_sim and sim_pista != "TODAS":
                                    df_sim = df_sim[df_sim[col_pista_sim].astype(str).str.upper() == sim_pista]

                                col_ha = 'AREA_MAESTRA'
                                if col_ha in df_sim.columns:
                                    df_sim[col_ha] = pd.to_numeric(df_sim[col_ha].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
                                    df_sim = df_sim[df_sim[col_ha] > 0]

                                if df_sim.empty:
                                    st.warning("⚠️ No se encontraron Órdenes de Servicio para los parámetros seleccionados.")
                                else:
                                    import math
                                    def red_excel(num):
                                        return math.floor(num + 0.5) if num >= 0 else math.ceil(num - 0.5)

                                    # 🎯 FILTRO FRANCOTIRADOR: Identificar estrictamente las columnas unitarias correctas
                                    col_tarifa_avion = None
                                    col_dominical = None
                                    
                                    for c in df_sim.columns:
                                        c_upper = str(c).upper().strip()
                                        # Bloqueo absoluto a las columnas que inflan los datos
                                        if any(x in c_upper for x in ["ORDEN", "SERVICIO", "FACTURAR", "PRODUCTOR", "TOTAL"]):
                                            continue
                                        if "AVION" in c_upper or "AVIÓN" in c_upper:
                                            col_tarifa_avion = c
                                        if "DOMINIC" in c_upper:
                                            col_dominical = c

                                    col_os = next((c for c in df_sim.columns if "OS" in str(c).upper() and "COSTO" not in str(c).upper()), df_sim.columns[0])
                                    for c in df_sim.columns:
                                        if str(c).upper().strip() in ["OS", "ORDEN", "Nº OS", "Nº ORDEN", "ORDEN DE SERVICIO"]:
                                            col_os = c; break

                                    col_finca = 'FINCA_MAESTRA'
                                    matriz_simulacion = []

                                    for _, row in df_sim.iterrows():
                                        os_val = str(row[col_os]).strip()
                                        if os_val == "" or os_val == "nan" or (os_val.replace('.','').isdigit() and len(os_val) > 5 and not os_val.startswith('318') and not os_val.startswith('319')): 
                                            continue

                                        finca_val = str(row[col_finca]).upper().strip()
                                        ha_val = float(row[col_ha])
                                        pista_val = str(row[col_pista_sim]).upper().strip() if col_pista_sim else "N/A"
                                        
                                        if pd.notna(row['FECHA_DT']):
                                            fecha_val = row['FECHA_DT'].strftime('%d/%m/%Y')
                                            semana_val = (row['FECHA_DT'] + pd.Timedelta(days=2)).isocalendar()[1]
                                        else:
                                            fecha_val = str(row['FECHA_MAESTRA'])
                                            col_sem = next((c for c in df_sim.columns if "SEMANA" in str(c).upper()), None)
                                            semana_val = row[col_sem] if col_sem else "N/A"

                                        # 🎯 EXTRACCIÓN UNITARIA PURA (Columnas T y U)
                                        tar_avion_raw = convertir_pesos(row[col_tarifa_avion]) if col_tarifa_avion and col_tarifa_avion in row else 0.0
                                        tar_dom_raw = convertir_pesos(row[col_dominical]) if col_dominical and col_dominical in row else 0.0
                                        
                                        # CORRECCIÓN DE VARIABLE CRÍTICA UNIFICADA
                                        tarifa_unitaria_actual = tar_avion_raw + tar_dom_raw

                                        if tarifa_unitaria_actual > 0 and ha_val > 0:
                                            
                                            # 1. Tarifa Actual Redondeada (ej. 51.906)
                                            t_act_red = red_excel(tarifa_unitaria_actual)
                                            
                                            # 2. Ingeniería Inversa: Hallar Base y Proyectar Nueva Tarifa Unitaria
                                            base_neta_ha = tarifa_unitaria_actual / (1 + (margen_actual / 100))
                                            tarifa_nueva_unitaria = base_neta_ha * (1 + (margen_nuevo / 100))
                                            t_nue_red = red_excel(tarifa_nueva_unitaria)
                                            
                                            # 3. La Ciencia del Excel: Resta de Tarifas Redondeadas × Hectáreas
                                            resta_tarifas = t_nue_red - t_act_red
                                            diferencia_total = red_excel(resta_tarifas * ha_val)

                                            # 4. Totales resultantes para consistencia del informe
                                            total_actual = red_excel(t_act_red * ha_val)
                                            total_nuevo = red_excel(t_nue_red * ha_val)

                                            matriz_simulacion.append({
                                                "Nº OS": os_val,
                                                "FECHA": fecha_val,
                                                "SEMANA": int(semana_val) if str(semana_val).isdigit() else semana_val,
                                                "FINCA": finca_val,
                                                "PISTA": pista_val,
                                                "HECTÁREAS": ha_val,
                                                f"TARIFA ACTUAL / Ha ({margen_actual}%)": t_act_red,
                                                f"NUEVA TARIFA / Ha ({margen_nuevo}%)": t_nue_red,
                                                "TOTAL ACTUAL ($)": total_actual,
                                                "NUEVO TOTAL ($)": total_nuevo,
                                                "DIFERENCIA ($)": diferencia_total
                                            })

                                    if not matriz_simulacion:
                                        st.warning("⚠️ Error de lectura: El sistema no detectó las columnas unitarias de Tarifa de Avión o Dominical. Verifique el archivo de origen.")
                                    else:
                                        df_resultados = pd.DataFrame(matriz_simulacion)

                                        df_semanal = df_resultados.groupby("SEMANA").agg({
                                            "HECTÁREAS": "sum",
                                            "TOTAL ACTUAL ($)": "sum",
                                            "NUEVO TOTAL ($)": "sum",
                                            "DIFERENCIA ($)": "sum"
                                        }).reset_index()
                                        df_semanal = df_semanal.sort_values(by="SEMANA").reset_index(drop=True)

                                        total_actual_global = df_resultados["TOTAL ACTUAL ($)"].sum()
                                        total_simulado_global = df_resultados["NUEVO TOTAL ($)"].sum()
                                        total_diferencia_global = df_resultados["DIFERENCIA ($)"].sum()

                                        st.markdown("### 🎯 Impacto Financiero Real de la Simulación")
                                        k1, k2, k3 = st.columns(3)
                                        k1.metric(f"💰 Total Actual ({margen_actual}%)", f"$ {total_actual_global:,.0f}".replace(",", "."))
                                        k2.metric(f"📈 Proyección ({margen_nuevo}%)", f"$ {total_simulado_global:,.0f}".replace(",", "."))
                                        color_delta = "normal" if total_diferencia_global > 0 else "inverse"
                                        k3.metric("⚖️ Dinero Real en Juego", f"$ {abs(total_diferencia_global):,.0f}".replace(",", "."), delta=f"$ {total_diferencia_global:,.0f}".replace(",", "."), delta_color=color_delta)

                                        st.markdown("<br>", unsafe_allow_html=True)
                                        tab_resumen, tab_detalle = st.tabs(["📊 1. Resumen Macroeconómico", "📋 2. Desglose Quirúrgico"])
                                        
                                        with tab_resumen:
                                            st.markdown("#### Matriz Semanal (Corte Sáb a Vie)")
                                            df_sem_vista = df_semanal.copy()
                                            df_sem_vista["HECTÁREAS"] = df_sem_vista["HECTÁREAS"].map("{:,.2f}".format)
                                            for col in ["TOTAL ACTUAL ($)", "NUEVO TOTAL ($)", "DIFERENCIA ($)"]:
                                                df_sem_vista[col] = df_sem_vista[col].map("$ {:,.0f}".format).str.replace(",", ".")
                                            st.dataframe(df_sem_vista, use_container_width=True, hide_index=True)
                                        
                                        with tab_detalle:
                                            st.markdown("#### Historial por OS")
                                            df_vista = df_resultados.copy()
                                            df_vista["HECTÁREAS"] = df_vista["HECTÁREAS"].map("{:,.2f}".format)
                                            columnas_moneda = [f"TARIFA ACTUAL / Ha ({margen_actual}%)", f"NUEVA TARIFA / Ha ({margen_nuevo}%)", "TOTAL ACTUAL ($)", "NUEVO TOTAL ($)", "DIFERENCIA ($)"]
                                            for col in columnas_moneda:
                                                df_vista[col] = df_vista[col].map("$ {:,.0f}".format).str.replace(",", ".")
                                            
                                            def col_dif(val):
                                                if isinstance(val, str) and "-" in val: return 'color: #721c24; background-color: #f8d7da; font-weight: bold; text-align: center;'
                                                elif isinstance(val, str) and "$" in val: return 'color: #155724; background-color: #d4edda; font-weight: bold; text-align: center;'
                                                return ''
                                            st.dataframe(df_vista.style.map(col_dif, subset=["DIFERENCIA ($)"]), use_container_width=True, hide_index=True)

                                        import io
                                        buffer_neg = io.BytesIO()
                                        with pd.ExcelWriter(buffer_neg, engine='openpyxl') as writer:
                                            df_semanal.to_excel(writer, sheet_name='Resumen_Semanal', index=False)
                                            df_resultados.to_excel(writer, sheet_name='Detalle_OS', index=False)
                                            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                                            borde = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                                            fondo = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                                            blanca = Font(color="FFFFFF", bold=True)
                                            for name in ['Resumen_Semanal', 'Detalle_OS']:
                                                ws = writer.sheets[name]
                                                ws.sheet_view.showGridLines = True
                                                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                                                    for cell in row:
                                                        cell.border = borde
                                                        if cell.row == 1:
                                                            cell.fill = fondo; cell.font = blanca; cell.alignment = Alignment(horizontal='center', vertical='center')
                                                        else:
                                                            if "HECTÁREAS" in str(ws.cell(1, cell.column).value): cell.number_format = '#,##0.00'
                                                            elif "($" in str(ws.cell(1, cell.column).value) or "%" in str(ws.cell(1, cell.column).value): cell.number_format = '"$" #,##0'
                                                    for col in ws.columns:
                                                        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col) + 4, 32)
                                        
                                        st.markdown("<br>", unsafe_allow_html=True)
                                        st.download_button(
                                            label="📥 DESCARGAR INFORME DUAL (EXCEL OFICIAL)",
                                            data=buffer_neg.getvalue(),
                                            file_name=f"Auditoria_Tarifas_{sim_pista}_{meses_dict[sim_mes]}_{sim_anio}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            type="primary",
                                            use_container_width=True
                                        )
                                        
                                        
                    
                    else:
                        st.error("❌ **ERROR DE RADAR:** No se detectó la columna 'FECHA' unificada.")
                else:
                    st.error("❌ **ERROR DE ALINEACIÓN:** No se logró estandarizar Fincas y Costos. Revise encabezados.")
            else:
                st.error("❌ **ERROR DE VOLUMEN:** Uno de los archivos está vacío.")

        except Exception as e:
            st.error(f"🛰️ **FALLO EN LOS MOTORES:** Error crítico. Motivo: {str(e)}")
