import matplotlib
matplotlib.use('Agg')
import streamlit as st
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

import pandas as pd
from datetime import datetime
import gspread
import time
import base64
import os
import math
import traceback
from supabase import create_client, Client

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

# --- LOGIN ---
if not st.session_state['autenticado']:
    st.markdown("<style>[data-testid='stSidebar'] {display: none;}</style>", unsafe_allow_html=True)
    st.markdown("<br><br>", unsafe_allow_html=True)
    c_log1, c_log2, c_log3 = st.columns([1, 1.2, 1])
    with c_log2:
        st.markdown("<h2 style='text-align: center; color: #0d1b2a;'>🚀 GÉNESIS OMEGA PRO</h2>", unsafe_allow_html=True)
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

# --- CONEXIÓN A BÓVEDA ---
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

@st.cache_data(show_spinner=False, ttl=1800)
def descargar_matriz_rapida(url, pestaña):
    try:
        if "gcp_service_account" in st.secrets: sat = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        else: sat = gspread.service_account(filename='credenciales.json')
        hoja = next((s for s in sat.open_by_url(url).worksheets() if "TABLA 1" in s.title.upper()), sat.open_by_url(url).sheet1) if "TABLA 1" in pestaña.upper() else sat.open_by_url(url).worksheet(pestaña)
        return hoja.get_all_values(value_render_option='UNFORMATTED_VALUE')
    except Exception: return []

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.get('usuario_nombre')}**")
    st.markdown("---")
    
    if st.session_state.get('usuario_rol') == "ADMIN":
        # ⚡ EL BOTÓN GLOBAL AHORA LLAMA A LA FUNCIÓN DEL MÓDULO 0
        if st.button("🔄 Sincronizar Nube", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.cache_resource.clear()
            m0.ordenar_base_datos_global()
            
        st.radio("🛰️ SELECCIONE LA OPERACIÓN:", [
            "🏠 Centro de Mando", "🛠️ 1. Mantenimiento Plantilla SAP", "📥 2. Carga Facturación", "⚙️ 3. Validación de Misión", "⌨️ 4. Ingreso Manual Acelerado (OS)", "📈 5. Sincronización Precios", "✈️ 6. Rastreo Dominicales", "⚖️ 7. Arqueo de Inventarios", "📊 8. Reporte Hectáreas (Pistas)", "📈 9. Dashboard Táctico", "📊 10. Inteligencia de Costos (BI)", "📜 11. Manual de Gobierno Técnico", "🚁 12. Simulador Financiero Libre", "🔮 13. El Oráculo (Inventarios)", "💰 14. Pronóstico Financiero", "🗺️ 15. Mapa de Calor Agronómico", "💼 16. Comparativo Gerencial (Dron vs Avión)", "🚀 17. Mega-Proyección Operativa", "🔍 18. Auditoría y Desglose Financiero"
        ], key="modulo_actual")
    else: 
        st.info("🛰️ Modo Consulta Gerencial Activado.")
        st.radio("📊 SELECCIONE EL REPORTE:", ["📈 9. Dashboard Táctico", "📊 10. Inteligencia de Costos (BI)", "💼 16. Comparativo Gerencial (Dron vs Avión)"], key="modulo_actual")

    st.button("🔒 CERRAR SESIÓN", use_container_width=True, on_click=lambda: st.session_state.update(autenticado=False))

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
