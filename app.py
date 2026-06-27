import pandas as pd
import streamlit as st
from datetime import datetime
import gspread
import time
import base64
import os

# ⚙️ REGLA DE ORO: Configuración de página primero
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 🛰️ CONEXIÓN DE HANGARES MODULARES ---
from modulos.utilidades import purificar_lote, quitar_tildes, extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro, limpiar_val_dom, procesar_fecha_pesada
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
import modulos.m11_manual_tecnico as m11 
import modulos.m12_simulador_agro as m12
import modulos.m13_oraculo as m13 # <-- AGREGUE ESTA LÍNEA

# --- 🔐 CREDENCIALES DE BÓVEDA ---
USUARIOS_CREDENTIALS = {
    "usernames": {
        "comandante": {"name": "Comandante Omega", "password": st.secrets["passwords"]["comandante"] if "passwords" in st.secrets else "123", "role": "ADMIN"},
        "gerencia": {"name": "Visor Gerencial", "password": st.secrets["passwords"]["gerencia"] if "passwords" in st.secrets else "123", "role": "VIEWER"}
    }
}

# 💥 BLINDAJE DE MEMORIA: Se añaden las anclas de estado
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if 'usuario_rol' not in st.session_state: st.session_state['usuario_rol'] = None
if 'usuario_nombre' not in st.session_state: st.session_state['usuario_nombre'] = None
if 'modulo_actual' not in st.session_state: st.session_state['modulo_actual'] = "🏠 Centro de Mando"

try: 
    import matplotlib
    HAS_MATPLOTLIB = True
except ImportError: 
    HAS_MATPLOTLIB = False

# --- 🛡️ MOTOR DE MARCA DE AGUA FANTASMA (4%) ---
try:
    with open("escudo.png", "rb") as image_file:
        bg_image = f"data:image/png;base64,{base64.b64encode(image_file.read()).decode()}"
    st.markdown(f"""
    <style>
    .stApp::before {{
        content: ""; background-image: url('{bg_image}');
        background-size: 550px; background-repeat: no-repeat; background-position: center;
        opacity: 0.04; position: fixed; top: 0; left: 0; bottom: 0; right: 0; z-index: 0; pointer-events: none;
    }}
    </style>
    """, unsafe_allow_html=True)
except: pass

# --- 🛡️ ARTILLERÍA VISUAL Y CSS BLINDADO ---
st.markdown("""
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stAppDeployButton { display: none !important; }
.viewerBadge_container { display: none !important; visibility: hidden !important; opacity: 0 !important; }
div[class^="viewerBadge"] { display: none !important; }
footer { display: none !important; visibility: hidden !important; }
#MainMenu { visibility: visible !important; display: block !important; }

.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }

[data-testid="stSidebar"] input { color: #0d1b2a !important; background-color: #ffffff !important; }
[data-testid="stSidebar"] button svg { fill: #0d1b2a !important; color: #0d1b2a !important; }

[data-testid="stSidebar"] button[kind="secondary"] {
    background-color: #ef4444 !important; border: 2px solid #b91c1c !important; border-radius: 8px !important; color: #ffffff !important;
}
[data-testid="stSidebar"] button[kind="secondary"]:hover { background-color: #dc2626 !important; }
[data-testid="stSidebar"] button[kind="secondary"] p { color: #ffffff !important; }

button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }

.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase; position: relative; z-index: 1;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px; position: relative; z-index: 1;}
div[data-baseweb="input"] input, div[data-baseweb="select"] { color: black !important; background-color: white !important; font-weight: bold; }
th { background-color: #f0f2f6 !important; color: black !important; }
[data-testid="stVerticalBlock"] { position: relative; z-index: 1; }
div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="number"] > div { background-color: #ffffff !important; border: 2px solid #0d1b2a !important; box-shadow: 1px 1px 4px rgba(0,0,0,0.05) !important; }
</style>
""", unsafe_allow_html=True)

# --- 3. 🔐 CONTROL DE ACCESO CENTRALIZADO (LOGIN) ---
if not st.session_state['autenticado']:
    st.markdown("<style>[data-testid='stSidebar'] {display: none;}</style>", unsafe_allow_html=True)
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    c_log1, c_log2, c_log3 = st.columns([1, 1.2, 1])
    with c_log2:
        if os.path.exists("escudo.png"):
            try: st.image("escudo.png", use_container_width=True)
            except: st.markdown("<h1 style='text-align: center; color: #D97706; font-size: 5rem;'>🛡️</h1>", unsafe_allow_html=True)
        else:
            st.markdown("<h2 style='text-align: center; color: #0d1b2a;'>🚀 GÉNESIS OMEGA PRO</h2>", unsafe_allow_html=True)
            
        st.markdown("<h2 style='text-align: center; color: #0d1b2a; margin-top: 10px; font-weight: bold;'>GÉNESIS AGROAÉREO</h2>", unsafe_allow_html=True)
        
        with st.form("Formulario"):
            u_in = st.text_input("🛰️ Usuario:", placeholder="Ingrese su usuario")
            p_in = st.text_input("🔑 Contraseña:", type="password", placeholder="Ingrese su contraseña")
            if st.form_submit_button("🔓 ACTIVAR SISTEMA", use_container_width=True):
                if u_in in USUARIOS_CREDENTIALS["usernames"] and p_in == USUARIOS_CREDENTIALS["usernames"][u_in]["password"]:
                    st.session_state['autenticado'] = True
                    st.session_state['usuario_rol'] = USUARIOS_CREDENTIALS["usernames"][u_in]["role"]
                    st.session_state['usuario_nombre'] = USUARIOS_CREDENTIALS["usernames"][u_in]["name"]
                    st.rerun()
                else: 
                    st.error("🚨 Credenciales incorrectas.")
    st.stop() 

# --- 4. CONEXIÓN SATELITAL GLOBAL ---
@st.cache_resource(show_spinner=False)
def conectar_satelite():
    return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"])) if "gcp_credentials" in st.secrets else gspread.service_account(filename='credenciales.json')

@st.cache_data(show_spinner=False, ttl=1800)
def descargar_matriz_rapida(url, pestaña):
    for i in range(3):
        try:
            hoja = next((s for s in conectar_satelite().open_by_url(url).worksheets() if "TABLA 1" in s.title.upper()), conectar_satelite().open_by_url(url).sheet1) if "TABLA 1" in pestaña.upper() else conectar_satelite().open_by_url(url).worksheet(pestaña)
            return hoja.get_all_values(value_render_option='UNFORMATTED_VALUE')
        except:
            if i < 2: time.sleep(2); continue
            else: return []

# --- 5. MENÚ MAESTRO TÁCTICO ---
with st.sidebar:
    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
    with col_img2:
        try: 
            st.image("escudo.png", use_container_width=True)
        except: 
            st.markdown("<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
            
    st.markdown(f"<p style='text-align: center; color: white; font-size:14px; font-weight: bold;'>👤 {st.session_state['usuario_nombre']}</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    if st.session_state['usuario_rol'] == "ADMIN":
        if st.button("🔄 Cargar Cócteles / Aviones", type="primary", use_container_width=True): 
            st.cache_data.clear()
            st.rerun()
            
        # 💥 BLINDAJE DE MEMORIA: Se asigna el selector a la variable 'key' para que no se borre
        st.radio("🛰️ SELECCIONE LA OPERACIÓN:", [
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
            "📊 10. Inteligencia de Costos (BI)",
            "📜 11. Manual de Gobierno Técnico",
            "🚁 12. Simulador Financiero Libre",
            "🔮 13. El Oráculo (Inventarios)" # <-- AGREGUE ESTA LÍNEA
        ], key="modulo_actual")
    else: 
        st.session_state['modulo_actual'] = "📈 9. Dashboard Táctico"
        st.info("🛰️ Modo Consulta Gerencial Activado.")
        
    st.markdown("---")
    if st.button("🔒 CERRAR SESIÓN", use_container_width=True):
        st.session_state['autenticado'], st.session_state['usuario_rol'], st.session_state['usuario_nombre'] = False, None, None
        st.session_state['modulo_actual'] = "🏠 Centro de Mando"
        st.rerun()

# --- 6. DELEGACIÓN A ESCUADRONES ---
# 💥 BLINDAJE DE MEMORIA: Se lee directamente desde la sesión guardada
menu = st.session_state['modulo_actual']

if menu == "🏠 Centro de Mando": m0.renderizar()
elif menu == "🛠️ 1. Mantenimiento Plantilla SAP": m1.ejecutar(extraer_numero)
elif menu == "📥 2. Carga Facturación": m2.ejecutar(extraer_numero)
elif menu == "⚙️ 3. Validación de Misión": m3.ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada)
elif menu == "⌨️ 4. Ingreso Manual Acelerado (OS)": m4.ejecutar(extraer_numero, purificar_lote)
elif menu == "📈 5. Sincronización Precios": m5.ejecutar(extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro)
elif menu == "✈️ 6. Rastreo Dominicales": m6.ejecutar(procesar_fecha_pesada, limpiar_val_dom)
elif menu == "⚖️ 7. Arqueo de Inventarios": m7.ejecutar(quitar_tildes, purificar_lote)
elif menu == "📊 8. Reporte Hectáreas (Pistas)": m8.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada, HAS_MATPLOTLIB)
elif menu == "📈 9. Dashboard Táctico": m9.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada)
elif menu == "📊 10. Inteligencia de Costos (BI)": m10.ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero)
elif menu == "📜 11. Manual de Gobierno Técnico": m11.ejecutar() 
elif menu == "🚁 12. Simulador Financiero Libre": m12.ejecutar(procesar_fecha_pesada, extraer_numero)
elif menu == "🔮 13. El Oráculo (Inventarios)": m13.ejecutar(purificar_lote, extraer_numero) # <-- AGREGUE ESTA LÍNEA
