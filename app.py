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
import modulos.m6_rastreo_dominicales as m6
import modulos.m5_sincronizacion_precios as m5
import modulos.m4_ingreso_manual as m4

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
    m3.ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada)
                    
# =====================================================================
# ⌨️ 4. INGRESO MANUAL ACELERADO Y LEGALIZACIÓN (OS)
# =====================================================================
elif menu == "⌨️ 4. Ingreso Manual Acelerado (OS)":
    m4.ejecutar(extraer_numero, purificar_lote)
# =====================================================================
# 📈 5. SINCRONIZACIÓN PRECIOS Y TARIFARIO MAESTRO
# =====================================================================
elif menu == "📈 5. Sincronización Precios":
    m5.ejecutar(extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro)
            
# =====================================================================
# ✈️ 6. RASTREO DOMINICALES
# =====================================================================
elif menu == "✈️ 6. Rastreo Dominicales":
    m6.ejecutar(procesar_fecha_pesada, limpiar_val_dom)
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
