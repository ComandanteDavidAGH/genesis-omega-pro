import pandas as pd
import streamlit as st
import io
import json
import re
import unicodedata
from datetime import datetime
import dateutil.parser
import openpyxl
import gspread
import plotly.express as px
import plotly.graph_objects as go

# --- 🛰️ CONEXIÓN DE HANGARES MODULARES (ESCUADRONES) ---
import modulos.m0_centro_mando as m0
import modulos.m1_mantenimiento as m1
import modulos.m2_facturacion as m2
import modulos.m3_validacion_facturacion as m3
import modulos.m4_ingreso_manual as m4
import modulos.m5_sincronizacion_precios as m5
import modulos.m6_rastreo_dominicales as m6
import modulos.m7_arqueo_inventarios as m7
import modulos.m8_reporte_hectareas as m8
import modulos.m9_dashboard_tactico as m9
import modulos.m10_bi_tarifas as m10

# Aquí definimos quiénes tienen acceso al sistema extrayendo claves de la bóveda
USUARIOS_CREDENTIALS = {
    "usernames": {
        "comandante": {
            "name": "Comandante Omega",
            "password": st.secrets["passwords"]["comandante"] if "passwords" in st.secrets else "123",
            "role": "ADMIN"
        },
        "gerencia": {
            "name": "Visor Gerencial / Cliente",
            "password": st.secrets["passwords"]["gerencia"] if "passwords" in st.secrets else "123",
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

# --- 2. ARTILLERÍA VISUAL Y CSS (SIGILO + HAMBURGUESA PROTEGIDA) ---
st.markdown("""
<style>
/* 🛡️ DESTRUCCIÓN DE GITHUB Y DEPLOY (Pero dejando la barra superior viva) */
[data-testid="stToolbarActions"] { display: none !important; }
.stAppDeployButton { display: none !important; }

/* 🛡️ DESTRUCCIÓN DE LA MARCA DE AGUA (Manage App) Y FOOTER */
.viewerBadge_container { display: none !important; visibility: hidden !important; opacity: 0 !important; }
div[class^="viewerBadge"] { display: none !important; }
footer { display: none !important; visibility: hidden !important; }

/* 🛡️ PROTECCIÓN ABSOLUTA DE LA HAMBURGUESA */
#MainMenu { visibility: visible !important; display: block !important; }

/* Resto de la Artillería Visual */
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }
button[kind="secondary"] { background-color: transparent !important; color: #0d1b2a !important; border: 1px solid #0d1b2a !important; transition: 0.3s; }
button[kind="secondary"]:hover { background-color: #0d1b2a !important; color: #d4af37 !important; }
[data-testid="stSidebar"] button[kind="secondary"] { color: white !important; border: 1px solid #d4af37 !important; }
[data-testid="stSidebar"] button[kind="secondary"]:hover { background-color: #d4af37 !important; color: #0d1b2a !important; }
div[data-baseweb="input"] input, div[data-baseweb="select"] { color: black !important; background-color: white !important; font-weight: bold; }
th { background-color: #f0f2f6 !important; color: black !important; }
</style>
""", unsafe_allow_html=True)
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
    if "gcp_credentials" in st.secrets:
        return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
    else:
        return gspread.service_account(filename='credenciales.json')

@st.cache_data(show_spinner=False, ttl=1800)
def descargar_matriz_rapida(url, pestaña):
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
    
    if st.button("🔒 Cerrar Sesión", use_container_width=True):
        st.session_state['autenticado'] = False
        st.session_state['usuario_rol'] = None
        st.session_state['usuario_nombre'] = None
        st.rerun()
        
    st.markdown("---")
    
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
        menu = "📈 9. Dashboard Táctico"
        st.info("🛰️ Modo Consulta Gerencial Activado. Acceso restringido a reportes visuales.")

    st.info(f"📅 Operación: {datetime.now().strftime('%Y-%m-%d')}")
    

# =====================================================================
# --- LLAMADOS A LOS ESCUADRONES (MÓDULOS) ---
# =====================================================================
if menu == "🏠 Centro de Mando":
    m0.renderizar()
elif menu == "🛠️ 1. Mantenimiento Plantilla SAP":
    m1.ejecutar(extraer_numero)
elif menu == "📥 2. Carga Facturación":
    m2.ejecutar(extraer_numero)
elif menu == "⚙️ 3. Validación de Misión":
    m3.ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada)
elif menu == "⌨️ 4. Ingreso Manual Acelerado (OS)":
    m4.ejecutar(extraer_numero, purificar_lote)
elif menu == "📈 5. Sincronización Precios":
    m5.ejecutar(extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro)
elif menu == "✈️ 6. Rastreo Dominicales":
    m6.ejecutar(procesar_fecha_pesada, limpiar_val_dom)
elif menu == "⚖️ 7. Arqueo de Inventarios":
    m7.ejecutar(quitar_tildes, purificar_lote)
elif menu == "📊 8. Reporte Hectáreas (Pistas)":
    m8.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada, HAS_MATPLOTLIB)
elif menu == "📈 9. Dashboard Táctico":
    m9.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada)
elif menu == "📊 10. Inteligencia de Costos (BI)":
    m10.ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero)
