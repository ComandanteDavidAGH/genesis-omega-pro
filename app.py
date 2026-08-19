# 1. 🛡️ SILENCIADOR GRÁFICO
import matplotlib
matplotlib.use('Agg')

# 2. ⚙️ INICIO DEL MOTOR STREAMLIT Y CONFIGURACIÓN
import streamlit as st
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# 3. 📦 LIBRERÍAS
import pandas as pd
from datetime import datetime
import gspread
import time
import base64
import os
import math
import traceback
from supabase import create_client, Client

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
import modulos.m13_oraculo as m13
import modulos.m14_presupuesto as m14
import modulos.m15_mapa_calor as m15
import modulos.m16_gerencia as m16
import modulos.m17_mega_proyeccion as m17
import modulos.m18_desglose_facturacion as m18
import modulos.m19_ingresos as m19
# 👇 INYECTA ESTA LÍNEA AQUÍ 👇
import modulos.m20_duelo_fincas as m20

# --- 🔐 CREDENCIALES ---
pass_cmd = st.secrets.get("passwords", {}).get("comandante", "123") if "passwords" in st.secrets else "123"
pass_ger = st.secrets.get("passwords", {}).get("gerencia", "123") if "passwords" in st.secrets else "123"

USUARIOS_CREDENTIALS = {
    "usernames": {
        "comandante": {"name": "Comandante Omega", "password": pass_cmd, "role": "ADMIN"},
        "gerencia": {"name": "Visor Gerencial", "password": pass_ger, "role": "VIEWER"}
    }
}

if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if 'usuario_rol' not in st.session_state: st.session_state['usuario_rol'] = None
if 'usuario_nombre' not in st.session_state: st.session_state['usuario_nombre'] = None
if 'modulo_actual' not in st.session_state: st.session_state['modulo_actual'] = "🏠 Centro de Mando"

HAS_MATPLOTLIB = True

# --- 🛡️ MARCA DE AGUA FANTASMA ---
try:
    if os.path.exists("escudo.png"):
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
except Exception: pass

# --- 🎯 ESTILIZACIÓN UNIFICADA TOTAL (AZUL MARINO & DORADO) ---
st.markdown("""
<style>
/* Ocultar elementos nativos innecesarios */
[data-testid="stToolbarActions"], .stAppDeployButton, .viewerBadge_container, div[class^="viewerBadge"], footer { display: none !important; }
#MainMenu { visibility: visible !important; display: block !important; }

.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }

/* =====================================================================
   🥇 REGLA MAESTRA UNIFICADA PARA TODOS LOS BOTONES DEL SISTEMA
   ===================================================================== */
div.stButton > button,
button[kind="primary"],
button[kind="secondary"],
[data-testid="stSidebar"] button {
    background-color: #0d1b2a !important;
    color: #d4af37 !important;
    border: 2px solid #d4af37 !important;
    border-radius: 8px !important;
    font-weight: 900 !important;
    box-shadow: 0px 4px 6px rgba(0,0,0,0.15) !important;
    transition: all 0.3s ease !important;
}

/* Efecto al pasar el mouse por encima (Hover) */
div.stButton > button:hover,
button[kind="primary"]:hover,
button[kind="secondary"]:hover,
[data-testid="stSidebar"] button:hover {
    background-color: #15283c !important;
    border-color: #f1c40f !important;
    box-shadow: 0px 0px 10px rgba(212, 175, 55, 0.6) !important;
}

/* Asegurar que el texto interno de los botones siempre sea Dorado */
div.stButton > button *,
button[kind="primary"] *,
button[kind="secondary"] *,
[data-testid="stSidebar"] button * {
    color: #d4af37 !important;
    font-weight: 900 !important;
}

/* =====================================================================
   📦 CASILLAS DE ENTRADA (INPUTS Y SELECTS)
   ===================================================================== */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
input[type="text"], 
input[type="password"] {
    border: none !important;
    outline: none !important;
    box-shadow: none !important;
    color: #0d1b2a !important;
    font-weight: 900 !important;
    font-size: 15px !important;
    background: transparent !important;
    padding: 8px 12px !important;
}

div[data-testid="stTextInput"] > div,
div[data-testid="stNumberInput"] > div,
div[data-baseweb="select"] {
    border: 2px solid #0d1b2a !important;
    border-radius: 8px !important;
    background-color: #ffffff !important;
    box-shadow: 0px 2px 5px rgba(0,0,0,0.05) !important;
    overflow: hidden !important;
}

div[data-testid="stTextInput"] > div:focus-within,
div[data-testid="stNumberInput"] > div:focus-within,
div[data-baseweb="select"]:focus-within {
    border: 2px solid #d4af37 !important;
    box-shadow: 0px 0px 8px rgba(212, 175, 55, 0.8) !important;
}

.titulo-principal { 
    color: #0d1b2a; 
    font-family: 'Arial Black', sans-serif; 
    border-bottom: 3px solid #d4af37; 
    text-transform: uppercase; 
    position: relative; 
    z-index: 1;
}
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
            except Exception: st.markdown("<h1 style='text-align: center; color: #D97706; font-size: 5rem;'>🛡️</h1>", unsafe_allow_html=True)
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

# --- CONEXIONES ---
@st.cache_resource(show_spinner=False)
def conectar_satelite():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        elif "gcp_credentials" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception: return None

@st.cache_data(show_spinner=False, ttl=1800)
def descargar_matriz_rapida(url, pestaña):
    sat = conectar_satelite()
    if not sat: return []
    for i in range(3):
        try:
            hoja = next((s for s in sat.open_by_url(url).worksheets() if "TABLA 1" in s.title.upper()), sat.open_by_url(url).sheet1) if "TABLA 1" in pestaña.upper() else sat.open_by_url(url).worksheet(pestaña)
            return hoja.get_all_values(value_render_option='UNFORMATTED_VALUE')
        except Exception:
            if i < 2: time.sleep(2); continue
            else: return []

@st.cache_resource(show_spinner=False)
def conectar_supabase():
    try:
        url = st.secrets.get("SUPABASE_URL") or st.secrets.get("supabase", {}).get("url")
        key = st.secrets.get("SUPABASE_KEY") or st.secrets.get("supabase", {}).get("key")
        if url and key: return create_client(url, key)
    except Exception: pass
    return None

supabase_client = conectar_supabase()
st.session_state['supabase'] = supabase_client

# --- MENÚ LATERAL ---
with st.sidebar:
    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
    with col_img2:
        try: 
            if os.path.exists("escudo.png"): st.image("escudo.png", use_container_width=True)
            else: st.markdown("<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
        except Exception: st.markdown("<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
            
    st.markdown(f"<p style='text-align: center; color: white; font-size:14px; font-weight: bold;'>👤 {st.session_state.get('usuario_nombre', 'Comandante')}</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    if st.session_state.get('usuario_rol') == "ADMIN":
        if st.button("🔄 Sincronizar Nube", use_container_width=True):
            st.cache_data.clear()
            st.cache_resource.clear()
            # 🛡️ Escudo protector inyectado:
            try:
                m0.ordenar_base_datos_global()
                st.success("✅ Nube sincronizada con éxito.")
            except Exception as e:
                st.error(f"🚨 Error en sincronización: {e}")
                
        # 💥 AQUÍ ESTABA EL ERROR: Le faltaba el key="modulo_actual" al final
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
            "🔮 13. El Oráculo (Inventarios)",
            "💰 14. Pronóstico Financiero",
            "🗺️ 15. Mapa de Calor Agronómico",
            "💼 16. Comparativo Gerencial (Dron vs Avión)",
            "🚀 17. Mega-Proyección Operativa",  
            "🔍 18. Auditoría y Desglose Financiero",
            "📦 19. Control y Auditoría de Ingresos",
            "⚔️ 20. Duelo de Titanes (Finca vs Finca)"
        ], key="modulo_actual") # 👈 ¡ESTA ES LA LLAVE MAESTRA!
    else: 
        st.info("🛰️ Modo Consulta Gerencial Activado.")
        st.radio("📊 SELECCIONE EL REPORTE:", [
            "📈 9. Dashboard Táctico", 
            "📊 10. Inteligencia de Costos (BI)",
            "💼 16. Comparativo Gerencial (Dron vs Avión)"
        ], key="modulo_actual")
        
    st.markdown("---")
    
    def apagar_motores():
        st.session_state['autenticado'] = False
        st.session_state['usuario_rol'] = None
        st.session_state['usuario_nombre'] = None
        st.session_state['modulo_actual'] = "🏠 Centro de Mando"

    st.button("🔒 CERRAR SESIÓN", use_container_width=True, on_click=apagar_motores)

# --- RUTEO DE MÓDULOS ---
menu = st.session_state.get('modulo_actual', "🏠 Centro de Mando")

if menu == "🏠 Centro de Mando": m0.renderizar()
elif menu == "🛠️ 1. Mantenimiento Plantilla SAP": m1.ejecutar(extraer_numero)
elif menu == "📥 2. Carga Facturación": m2.ejecutar(extraer_numero)
elif menu == "⚙️ 3. Validación de Misión": m3.ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada)
elif menu == "⌨️ 4. Ingreso Manual Acelerado (OS)": m4.ejecutar(extraer_numero, purificar_lote)
elif menu == "📈 5. Sincronización Precios": m5.ejecutar(supabase_client, extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro)
elif menu == "✈️ 6. Rastreo Dominicales": m6.ejecutar(procesar_fecha_pesada, limpiar_val_dom)
elif menu == "⚖️ 7. Arqueo de Inventarios": m7.ejecutar(quitar_tildes, purificar_lote)
elif "8. Reporte Hectáreas (Pistas)" in menu: m8.ejecutar(supabase_client, descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada)
elif menu == "📈 9. Dashboard Táctico": m9.ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada)
elif menu == "📊 10. Inteligencia de Costos (BI)": m10.ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero)
elif menu == "📜 11. Manual de Gobierno Técnico": m11.ejecutar() 
elif menu == "🚁 12. Simulador Financiero Libre": m12.ejecutar(procesar_fecha_pesada, extraer_numero)
elif menu == "🔮 13. El Oráculo (Inventarios)": m13.ejecutar(purificar_lote, extraer_numero)
elif menu == "💰 14. Pronóstico Financiero": m14.ejecutar(purificar_lote, extraer_numero)
elif menu == "🗺️ 15. Mapa de Calor Agronómico": m15.ejecutar(purificar_lote, extraer_numero)
elif menu == "💼 16. Comparativo Gerencial (Dron vs Avión)": m16.ejecutar()
elif menu == "🚀 17. Mega-Proyección Operativa": m17.ejecutar(supabase_client)
elif menu == "🔍 18. Auditoría y Desglose Financiero": m18.ejecutar()
# 💥 AQUÍ INYECTAMOS LA LÓGICA DE DIRECCIONAMIENTO DEL MÓDULO 19:
elif menu == "📦 19. Control y Auditoría de Ingresos": m19.ejecutar()
# 👇 INYECTA ESTAS LÍNEAS AL FINAL 👇
elif menu == "⚔️ 20. Duelo de Titanes (Finca vs Finca)": 
    m20.ejecutar()    
