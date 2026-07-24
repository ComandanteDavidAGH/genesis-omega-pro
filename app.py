# 1. 🛡️ SILENCIADOR GRÁFICO (DEBE IR ANTES QUE CUALQUIER OTRA COSA)
import matplotlib
matplotlib.use('Agg')

# 2. ⚙️ INICIO DEL MOTOR STREAMLIT Y CONFIGURACIÓN (REGLA DE ORO)
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

# 💥 ESCUDO GENERAL ANTI-PANTALLA BLANCA (Captura cualquier fallo y lo muestra en pantalla)
try:
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

    # --- 🔐 CREDENCIALES SEGURAS ---
    pass_cmd = "123"
    pass_ger = "123"
    if "passwords" in st.secrets:
        pass_cmd = st.secrets["passwords"].get("comandante", "123")
        pass_ger = st.secrets["passwords"].get("gerencia", "123")

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

    # --- 🛡️ MOTOR DE MARCA DE AGUA FANTASMA (4%) ---
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

    # --- 🎯 ARTILLERÍA VISUAL CENTRALIZADA ---
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

    th { background-color: #f0f2f6 !important; color: black !important; }
    [data-testid="stVerticalBlock"] { position: relative; z-index: 1; }

    div[data-testid="stMainBlockContainer"] div[data-testid="stTextInput"] input,
    div[data-testid="stMainBlockContainer"] div[data-testid="stSelectbox"] [data-baseweb="select"],
    div[data-testid="stMainBlockContainer"] div[data-testid="stNumberInput"] input,
    div[data-testid="stMainBlockContainer"] div[data-testid="stDateInput"] input {
        border: 2px solid #0d1b2a !important;
        background-color: #ffffff !important;
        border-radius: 8px !important;
        color: #0d1b2a !important;
        font-weight: 900 !important;
        font-size: 15px !important;
    }

    div[data-testid="stMainBlockContainer"] div[data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
        border: 2px dashed #0d1b2a !important;
        border-radius: 8px !important;
    }

    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"],
    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"] pre,
    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"] pre code {
        background-color: #ffffff !important;
        border: 3px solid #0d1b2a !important;
        border-radius: 8px !important;
    }
    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"] code,
    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"] code span,
    div[data-testid="stMainBlockContainer"] div[data-testid="stCodeBlock"] pre span {
        color: #0d1b2a !important;
        font-weight: 900 !important;
        font-size: 16px !important;
        font-family: 'Arial Black', monospace !important;
    }

    .stTextInput input,
    .stSelectbox span,
    .stNumberInput input,
    .stDateInput input {
        color: #000000 !important;
        font-weight: 900 !important;
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
                try: st.image("escudo.png")
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

    # --- 4. 🛰️ HUB DE CONEXIONES GLOBALES ---
    @st.cache_resource(show_spinner=False)
    def conectar_satelite():
        try:
            if "gcp_service_account" in st.secrets:
                return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
            elif "gcp_credentials" in st.secrets:
                return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
            return gspread.service_account(filename='credenciales.json')
        except Exception:
            return None

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

    def sincronizar_y_ordenar_tabla1_a_supabase():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        return

    def convertir_fecha_excel(val):
        if pd.isna(val) or val is None or str(val).strip() == "": return ""
        val_str = str(val).strip()
        if val_str.replace('.', '').isdigit():
            try:
                num = float(val_str)
                if 30000 < num < 60000:
                    return pd.to_datetime(num, unit='D', origin='1899-12-30').strftime('%d/%m/%Y')
            except Exception: pass
        return val_str

    def sanitizar(v):
        if pd.isna(v) or v is None: return ""
        if isinstance(v, (float, int)):
            if math.isnan(v) or math.isinf(v): return 0
            return v
        return str(v).strip()

    try:
        supabase = st.session_state['supabase']
        gc = conectar_satelite()
        if not gc: return

        with st.spinner("🔄 Rescatando Sábana con formato correcto de fechas..."):
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            t1_formulas = ws_t1.get_all_values(value_render_option='FORMULA')
            t1_valores = ws_t1.get_all_values(value_render_option='FORMATTED_VALUE')
            
            idx_header = 4
            for i in range(min(10, len(t1_formulas))):
                if "FINCA" in [str(x).upper().strip() for x in t1_formulas[i]]:
                    idx_header = i
                    break
                    
            headers_exactos = [str(x).replace('\n', ' ').strip() for x in t1_valores[idx_header]]
            num_cols = len(headers_exactos)
            
            datos_form = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_formulas[idx_header+1:]]
            datos_val = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_valores[idx_header+1:]]
            
            df_form = pd.DataFrame(datos_form, columns=headers_exactos)
            df_val = pd.DataFrame(datos_val, columns=headers_exactos)
            
            col_id = headers_exactos[0]
            filas_validas = df_val[col_id].astype(str).str.strip() != ""
            df_form = df_form[filas_validas].copy()
            df_val = df_val[filas_validas].copy()

            col_fecha = next((c for c in headers_exactos if "FECHA" in c.upper()), None)

            if col_fecha:
                df_val[col_fecha] = df_val[col_fecha].apply(convertir_fecha_excel)
                df_val['fecha_dt'] = pd.to_datetime(df_val[col_fecha], format='%d/%m/%Y', errors='coerce')
                
                df_val_sorted = df_val.sort_values(by='fecha_dt', ascending=False, na_position='last')
                indices_ord = df_val_sorted.index
                df_form_sorted = df_form.loc[indices_ord]
                df_val_sorted = df_val_sorted.drop(columns=['fecha_dt'])
            else:
                df_val_sorted = df_val
                df_form_sorted = df_form

            registros = []
            for _, row in df_val_sorted.iterrows():
                rec = {c: sanitizar(row[c]) for c in headers_exactos if c != ""}
                registros.append(rec)

            if registros:
                supabase.table("TABLA_1").delete().neq(col_id, "_VACIO_IMPOSIBLE_999_").execute()
                tamano_bloque = 250
                for i in range(0, len(registros), tamano_bloque):
                    supabase.table("TABLA_1").insert(registros[i:i + tamano_bloque]).execute()
                st.toast(f"⚡ Supabase sincronizado ({len(registros)} filas sin NULLs).", icon="✅")

            valores_drive = df_form_sorted[headers_exactos].fillna("").values.tolist()
            if valores_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                ws_t1.batch_clear([rango_borrar])
                ws_t1.update(range_name=rango_inicio, values=valores_drive, value_input_option='USER_ENTERED')
                st.toast("⚡ Google Drive Físicamente Ordenado.", icon="✅")

    except Exception as e:
        st.toast(f"🚨 Sincronización en proceso: {e}")

    @st.cache_resource(show_spinner=False)
    def conectar_supabase():
        try:
            url, key = None, None
            if "supabase" in st.secrets:
                url = st.secrets["supabase"].get("url")
                key = st.secrets["supabase"].get("key")
            if not url:
                url = st.secrets.get("SUPABASE_URL")
                key = st.secrets.get("SUPABASE_KEY")
            if url and key:
                return create_client(url, key)
        except Exception:
            pass
        return None

    try:
        supabase_client = conectar_supabase()
        st.session_state['supabase'] = supabase_client
    except Exception:
        supabase_client = None

    # --- 5. MENÚ MAESTRO TÁCTICO ---
    with st.sidebar:
        col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
        with col_img2:
            try: 
                if os.path.exists("escudo.png"):
                    st.image("escudo.png")
                else:
                    st.markdown("<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
            except Exception: 
                st.markdown("<h3 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h3>", unsafe_allow_html=True)
                
        st.markdown(f"<p style='text-align: center; color: white; font-size:14px; font-weight: bold;'>👤 {st.session_state.get('usuario_nombre', 'Comandante')}</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        if st.session_state.get('usuario_rol') == "ADMIN":
            if st.button("🔄 Sincronizar Nube", type="primary", use_container_width=True):
                st.cache_data.clear()
                st.cache_resource.clear()
                sincronizar_y_ordenar_tabla1_a_supabase()
                st.success("✅ Memoria purgada y sincronizada.")
                time.sleep(0.5)
                st.rerun()
                
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
                "🔍 18. Auditoría y Desglose Financiero"
                
            ], key="modulo_actual")
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

    # --- 6. DELEGACIÓN A ESCUADRONES ---
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

except Exception as e:
    st.error(f"🚨 ERROR CRÍTICO DETECTADO:\n{e}")
    st.code(traceback.format_exc(), language="python")
