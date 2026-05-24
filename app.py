import pandas as pd
import streamlit as st
import io
import json
import re
import unicodedata
from datetime import datetime
import dateutil.parser

# --- 🔐 BÓVEDA DE SEGURIDAD OMEGA ---
# Aquí definimos quiénes tienen acceso al sistema
USUARIOS_CREDENTIALS = {
    "usernames": {
        "comandante": {
            "name": "Comandante Omega",
            "password": "Alfa123*", # 🚨 CAMBIE ESTA CLAVE DESPUÉS
            "role": "ADMIN"
        },
        "gerencia": {
            "name": "Visor Gerencial / Cliente",
            "password": "Omega456*", # 🚨 CAMBIE ESTA CLAVE DESPUÉS
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
    st.markdown("<h1 class='titulo-principal'>Centro de Mando Omega Pro</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3>Bienvenido Comandante al Sistema Unificado:</h3>
        <p>Seleccione en el menú lateral la operación que desea realizar hoy. Los módulos están protegidos y operan de forma independiente.</p>
        <ol>
            <li><b>Mantenimiento:</b> Purifique y suba la Sábana SAP a la Bóveda (Plantilla).</li>
            <li><b>Facturación:</b> Cargue la sábana de SAP y los pedidos. Luego valide y facture en el módulo 3.</li>
            <li><b>Ingreso Manual Acelerado:</b> Digite los datos base de sus OS y el sistema calculará e inyectará el resto.</li>
            <li><b>Sincronización:</b> Actualice precios semanalmente simulando la Macro de VBA.</li>
            <li><b>Dominicales:</b> Rastree fechas de operación y recargos con inyección directa.</li>
            <li><b>Arqueo:</b> Auditoría total de pistas contra saldos SAP, con conciliación inteligente.</li>
            <li><b>Radar Hectáreas:</b> Visor dinámico semana a semana y mes a mes para gerencia.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

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
                    # Aquí la mira telescópica ubica la columna "Material"
                    col_mat = [c for c in df_p.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper()][0]
                    
                    finca_sap = str(match_sap.iloc[0][col_finca]).strip().upper()
                    
                    # 🎯 REGLA DE ORO 459: Francotirador directo a la columna Material
                    ha_correcta = 0.0
                    for _, fila_ped in match_sap.iterrows():
                        valor_material = str(fila_ped[col_mat]).strip()
                        # Si la columna Material es exactamente 459, extraemos las Hectáreas
                        if valor_material == "459" or valor_material.split(".")[0] == "459": 
                            ha_correcta = extraer_numero(fila_ped[col_ha])
                            break
                    
                    # Asignamos la cantidad encontrada
                    if ha_correcta > 0:
                        st.session_state['ha_radar_sap'] = ha_correcta
                    else:
                        st.session_state['ha_radar_sap'] = extraer_numero(match_sap.iloc[0][col_ha])
                    
                    st.success(f"✅ **SAP CONFIRMADO:** {finca_sap} | {st.session_state['ha_radar_sap']} Ha")
                except:
                    pass
        # ---------------------------
        # ---------------------------

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
# =======================================================
        # --- 1. RECONEXIÓN DE BASES DE DATOS ---
        # =======================================================
        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        # =======================================================
        # --- 2. IDENTIFICACIÓN DE LA FINCA ---
        # =======================================================
        import re 
        finca_limpia = re.sub(r'\s+', ' ', str(finca_sel)).strip().upper()

        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        
        # Primero buscamos qué tipo de productor es en la configuración general (df_t2)
        if not df_t2.empty:
            match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia]
            if not match_t2.empty:
                fila_t2 = match_t2.iloc[0]
                tipo_productor = str(fila_t2.iloc[5]).strip().upper()
                tipo_de_tope_finca = str(fila_t2.iloc[6]).strip().upper()
        
        # =======================================================
        # --- 3. EXTRACCIÓN DE TARIFAS ---
        # =======================================================
        # Ahora sí, con el tipo_productor claro, buscamos sus tarifas
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
        datos_raw = datos_vuelo['DATOS_FILA']
        
        num_pedido = "S/N"
        datos_vuelo = vuelos_informe[vuelos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo.get('DATOS_FILA', {})
        
        # --- 🎯 ENLACE MAESTRO DE PEDIDO SAP ---
        num_pedido = "S/N"
        
        # 1. Prioridad Máxima: Lo que el Comandante digitó en la casilla Radar (Módulo 3)
        if pedido_sap and len(str(pedido_sap)) >= 7:
            num_pedido = str(pedido_sap).strip()
            
        # 2. Segunda Prioridad: Lo que el Escáner automático capturó (Módulo 2)
        elif datos_vuelo.get('PEDIDO_SAP') and str(datos_vuelo.get('PEDIDO_SAP')).strip() != "":
            num_pedido = str(datos_vuelo.get('PEDIDO_SAP')).strip()
            
        # 3. Plan de Contingencia: Buscar en las columnas crudas del Excel (Legado)
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
            # --- TOMA DE DECISIÓN DE HECTÁREAS ---
            ha_sugerida = float(st.session_state.get('ha_radar_sap', 0.0))
            if ha_sugerida == 0.0:  # Si no hay datos de SAP, usamos el reporte del piloto
                ha_sugerida = float(ha_dosis_detectada)
                
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

        dict_topes_pista = {"TOPE MAX GENERAL": {"PLUC": 63325, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325}, "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517}, "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}}
        val_tope = dict_topes_pista.get(tipo_de_tope_finca, {}).get(pista_sel, 999999)
        
        # 🎯 AJUSTE DE PRECIOS EXACTOS SEGÚN IMAGEN MAESTRA (Flota Actualizada)
        dict_aviones = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA 25": 3036525, "AIR TRACTOR": 4665109, "CESSNA ASA": 3666600, "CESSNA FUMIGARAY": 3065952}
        dict_drones = {"DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280}

        with st.container(border=True):
            st.markdown("#### ✈️ Hangar de Despliegue")
            costo_total_vuelos = 0.0
            total_ha_cobro_escuadron = 0.0
            horometro_final_avion = 0.0  # 🎯 CAJA FUERTE PARA EL TIEMPO

            if mision_solo_dron:
                st.success("🚁 Modo Dron Activo: Costos calculados sin recargos terrestres ni topes de pista.")
                df_drones_def = pd.DataFrame([{"Drone": "DRONE DATAROT", "Hectáreas": float(ha_cobro_detectada)}])
                escuadron_drones = st.data_editor(df_drones_def, key=f"drones_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    costo_total_vuelos += (dict_drones.get(dr_sel, 0) * ha_dr) * mult_avion_final

            else:
                c_av, c_dr = st.columns(2)
                # 📡 RADAR DINÁMICO DE FLOTA (Ajustado a la estructura real de su Excel)
                try:
                    # 🚨 REEMPLACE 'SU_VARIABLE' POR LA QUE LEE LA PESTAÑA DE VALIDACIÓN
                    df_flota = st.session_state['SU_VARIABLE'] 
                    
                    # 🛩️ Extraer Aviones (Nombre en 'TIPO', Tarifa en 'HORA')
                    df_av = df_flota[df_flota['TIPO'].notna() & (df_flota['TIPO'].astype(str).str.strip() != '')]
                    # Quitamos los puntos de miles (Ej: 4.606.562 -> 4606562) para que Python pueda hacer la matemática
                    dict_aviones = dict(zip(df_av['TIPO'].astype(str).str.strip(), pd.to_numeric(df_av['HORA'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)))
                    
                    # 🚁 Extraer Drones (Nombre en 'Nombre Drone', Tarifa en 'Precio po...')
                    col_precio_dr = next((c for c in df_flota.columns if 'Precio po' in str(c)), 'Precio')
                    df_dr = df_flota[df_flota['Nombre Drone'].notna() & (df_flota['Nombre Drone'].astype(str).str.strip() != '')]
                    dict_drones = dict(zip(df_dr['Nombre Drone'].astype(str).str.strip(), pd.to_numeric(df_dr[col_precio_dr].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)))
                except Exception as e:
                    pass # Falla silenciosa anticaídas por si la tabla aún no ha cargado

                # 👇 COORDENADAS DE PANTALLA (Alineación perfecta de espacios)
                # 🚨 IMPORTANTE: Si su código usaba 'with col1:' en vez de 'with c_av:', cámbielo aquí abajo
                with c_av: 
                    st.markdown("##### 🛩️ Base Aviones")
                    df_aviones_def = pd.DataFrame([{"Avión": "THRUS SR2", "Hectáreas": float(ha_cobro_detectada), "Horómetro": 1.00}])
                    
                    # Si el radar falla, dejamos unos de reserva para que no se caiga la app
                    opciones_av = list(dict_aviones.keys()) if 'dict_aviones' in locals() and dict_aviones else ["THRUS SR2", "PIPER PA 36-375"]
                    
                    escuadron_aviones = st.data_editor(df_aviones_def, key=f"aviones_{casilla_key}", num_rows="dynamic", column_config={"Avión": st.column_config.SelectboxColumn("Modelo", options=opciones_av, required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f"), "Horómetro": st.column_config.NumberColumn("Horómetro", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)
                    
                with c_dr:
                    st.markdown("##### 🚁 Base Drones (Apoyo)")
                    df_drones_def = pd.DataFrame([{"Drone": None, "Hectáreas": 0.0}])
                    
                    opciones_dr = list(dict_drones.keys()) if 'dict_drones' in locals() and dict_drones else ["DRONE DATAROT", "DRON GENESYS"]
                    
                    escuadron_drones = st.data_editor(df_drones_def, key=f"drones_mix_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=opciones_dr), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)                # 👇 ESTOS CÁLCULOS AHORA VIVEN PROTEGIDOS DENTRO DEL "ELSE"
                for index, row in escuadron_aviones.iterrows():
                    av_sel = row["Avión"]
                    try:
                        val_ha = row.get("Hectáreas", 0)
                        ha_av = float(val_ha) if val_ha not in [None, "None", "", "nan"] else 0.0
                    except:
                        ha_av = 0.0
                        
                    try:
                        val_horo = row.get("Horómetro", 0)
                        horo = float(val_horo) if val_horo not in [None, "None", "", "nan"] else 0.0
                    except:
                        horo = 0.0
                    
                    if pd.isna(av_sel) or ha_av <= 0: continue
                    total_ha_cobro_escuadron += ha_av
                    horometro_final_avion += horo  # 🎯 ATRAPAMOS EL VALOR DE LA TABLA
                    tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av
                    tarifa_aplicada = tarifa_base_ha + recargo_final if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope) + recargo_final
                    costo_total_vuelos += (tarifa_aplicada * ha_av) * mult_avion_final
                    
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    costo_total_vuelos += (dict_drones.get(dr_sel, 0) * ha_dr) * mult_avion_final
            
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

                        # --- 🧮 CÁLCULOS MATEMÁTICOS DIRECTOS Y AJUSTE DE AERONAVE (Corregido) ---
                        ha_f = float(ha_dosis_final)
                        # 🎯 Inyecta el tiempo real digitado en la tabla superior
                        h_total_v = (ha_f / 10) if mision_solo_dron else horometro_final_avion
                        vol_total_gln = ha_f * 6
                        rend_min = h_total_v * 60
                        
                        # 🎯 CORRECCIÓN 1: Identificación Dinámica del Piloto/Operador
                        piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                        
                        # 🎯 CORRECCIÓN 2: Asignación e Inyección Dinámica de Matrículas (HK)
                        if mision_solo_dron:
                            if "DATAROT" in tipo_mision.upper(): hk_f = "DR51"
                            elif "GENESYS" in tipo_mision.upper(): hk_f = "DR52"
                            elif "AVIL" in tipo_mision.upper(): hk_f = "DR53"
                            else: hk_f = "DRONE_GEN"
                        else:
                            # Si es AVIÓN, toma la matrícula real seleccionada en los campos anteriores
                            hk_f = hk_sel if 'hk_sel' in locals() else "AVION_REG"

                        # 🎯 CORRECCIÓN 3: Modelo de la Aeronave para la base de datos
                        modelo_f = "DRONE" if mision_solo_dron else "AVION"

                        # 🚁 CÁLCULO DE PAGO A TERCEROS (Columna AD en Excel) - Ahora sí detecta el HK correcto
                        tarifa_pago = 84427 if "DR51" in hk_f else 71280 if ("DR52" in hk_f or "DR53" in hk_f) else 0
                        total_pago_avion = ha_f * tarifa_pago if mision_solo_dron else 0
                        try:
                            # Detecta cuántos datos hay guardados actualmente en la Tabla Azul
                            # para saber si la nueva factura va a caer en la fila 6, 7, 100, etc.
                            fila_excel = len(st.session_state['df_azul_actual']) + 2 
                        except:
                            fila_excel = 6 # Fila de respaldo por si el radar falla

                        
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
                        row_azul[10] = h_total_v
                        row_azul[11] = 6
                        row_azul[12] = round(vol_total_gln, 2)    # M: VOLUMEN (gln)
                        row_azul[13] = round(h_total_v, 2)        # N: RENDIMIENTO (hora)
                        row_azul[14] = round(rend_min, 2)         # O: RENDIMIENTO (min)
                        row_azul[15] = piloto_f
                        row_azul[16] = hk_f
                        row_azul[17] = tipo_mision
                        
                        # 💰 ZONA FINANCIERA CALIBRADA (RESCATE DE VARIABLE) 🎯🎯
                        # Como tarifa_aplicada se quedó en el radar anterior, deducimos el valor de la caja 429
                        # dividiendo el costo total de los vuelos entre las hectáreas reales.
                        tarifa_vuelo_plena = float(costo_total_vuelos / ha_f) if ha_f > 0 else 0.0
                        valor_dominical = float(recargo_final)
                        
                        row_azul[18] = float(gran_total)          # S: COSTO AVIÓN ($) [Total general de la factura]
                        
                        # PUNTO 1 y 2 COMPLETADOS: El 429 pleno RESTANDO el dominical (Valor Limpio del Vuelo a la Columna T)
                        row_azul[19] = round(tarifa_vuelo_plena - valor_dominical, 2)  # T: COSTO AVIÓN ($/ha) 
                        
                        # Aquí se almacena el dominical por separado para que no se duplique en SAP
                        row_azul[20] = round(valor_dominical, 2)  # U: DOMINIC ($/ha)
                        
                        row_azul[21] = float(gran_total)          # V: COSTO AVIÓN ($/finca)
                        
                        row_azul[23] = pista_manual               # X: PISTA
                        
                        row_azul[28] = float(gran_total)          # AC: COSTO TOTAL
                        row_azul[29] = float(total_pago_avion)    # AD: TOTAL PAGO AVIÓN 
                        row_azul[32] = tipo_productor             # AG: TIPO DE PRODUCTOR
                        row_azul[33] = "GÉNESIS_V2_PRO"           # AH: SELLO DE SISTEMA
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
                        
                        # 1. Escaneo de la Tabla Azul (Columna A: Nº de Orden)
                        col_azul = hoja_maestra.col_values(1)
                        fila_destino_azul = 1
                        for i in range(len(col_azul)-1, -1, -1):
                            if str(col_azul[i]).strip() != "":
                                fila_destino_azul = i + 2
                                break
                        
                        # 🚧 Si llegamos al límite físico de la hoja, construimos 10 filas más
                        if fila_destino_azul > hoja_maestra.row_count:
                            hoja_maestra.add_rows(10)

                        # 2. Escaneo de la Tabla Apoyo (Columna B: Finca)
                        col_apoyo = hoja_apoyo.col_values(2)
                        fila_destino_apoyo = 1
                        for i in range(len(col_apoyo)-1, -1, -1):
                            if str(col_apoyo[i]).strip() != "":
                                fila_destino_apoyo = i + 2
                                break
                                
                        # 🚧 Si llegamos al límite físico de la hoja de apoyo, construimos 10 filas más
                        if fila_destino_apoyo > hoja_apoyo.row_count:
                            hoja_apoyo.add_rows(10)

                        # 🔥 LA SOLUCIÓN TÁCTICA: Python calcula el número exacto y lo ancla como valor fijo
                        fila_apoyo[0] = fila_destino_apoyo - 3

                        # Inyectamos exactamente en las coordenadas calculadas sin chocar con el límite
                        hoja_maestra.update(range_name=f"A{fila_destino_azul}", values=[row_azul], value_input_option='USER_ENTERED')
                        hoja_apoyo.update(range_name=f"A{fila_destino_apoyo}", values=[fila_apoyo], value_input_option='USER_ENTERED')
                        # 🧪 DESEMBARCO DE QUÍMICOS EN LA PESTAÑA 'MEMORIA'
                        try:
                            filas_memoria = []
                            # Iteramos sobre la matriz de productos en pantalla
                            for idx, row in edited_df.iterrows():
                                nombre_prod = str(row.get("A: Producto", ""))
                                if "⚠️" not in nombre_prod and nombre_prod.strip() != "" and nombre_prod.lower() != "nan":
                                    dosis_prod = float(row.get("D: Dosis Total (Sistema)", 0))
                                    lote_prod = str(row.get("G: Lotes (SAP)", "S/N"))
                                    
                                    # Armamos la fila de 10 columnas según su diseño
                                    fila_m = [""] * 10
                                    fila_m[0] = fecha_str                                      # A: FECHA
                                    fila_m[1] = coctel_ganador                                 # B: ORDEN/COCTEL
                                    fila_m[2] = str(pista_manual).split("-")[0].strip()[:4]    # C: PISTA (Ej: LUCI, TEHO)
                                    fila_m[3] = nombre_prod                                    # D: PRODUCTO
                                    fila_m[4] = lote_prod                                      # E: LOTE
                                    fila_m[5] = float(dosis_prod)                              # F: CANTIDAD
                                    fila_m[6] = "BODEGA PRINCIPAL"                             # G: BODEGA
                                    fila_m[7] = ""                                             # H: MODELO (Vacío)
                                    fila_m[8] = "X"                                            # I: Facturado (X)
                                    fila_m[9] = finca_limpia                                   # J: FINCA
                                    
                                    filas_memoria.append(fila_m)
                                    
                            # Si recolectamos químicos, los disparamos todos juntos
                            if filas_memoria:
                                hoja_memoria.append_rows(filas_memoria, value_input_option='USER_ENTERED')
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
# 📈 5. SINCRONIZACIÓN PRECIOS
# =====================================================================
elif menu == "📈 5. Sincronización Precios":
    st.markdown("<h1 class='titulo-principal'>Sincronización Semanal de Precios</h1>", unsafe_allow_html=True)
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
elif menu == "⚖️ 7. Arqueo de Inventarios":
    
    st.markdown("<h1 class='titulo-principal'>⚖️ Arqueo de Inventarios y Conciliación</h1>", unsafe_allow_html=True)
    
    # 📦 ZONA DE CARGA EN PANTALLA PRINCIPAL (Descamuflada)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        archivo_sap = st.file_uploader("1️⃣ Sábana de SAP", type=['xlsx', 'csv'])
    with c2:
        st.markdown("### 📋 2. Reportes Físicos")
        archivos_sup = st.file_uploader("2️⃣ Reportes Supervisores (.xlsx)", type=['xlsx'], accept_multiple_files=True)
    with c3:
        st.markdown("### 🎯 3. Objetivo")
        semana_obj = st.text_input("Semana a Auditar (Ej: 17):", placeholder="Escriba la semana aquí...")

    if "arqueo_procesado" not in st.session_state:
        st.session_state.arqueo_procesado = False
    if "observaciones_memoria" not in st.session_state:
        st.session_state.observaciones_memoria = {}

    def generar_cruce():
        cruce = pd.merge(st.session_state.df_sap_grouped, st.session_state.df_sup_grouped, on=['PISTA', 'LOTE_KEY'], how='outer')
        
        cruce['ITEM'] = cruce['ITEM'].fillna("---")
        cruce['PRODUCTO'] = cruce['PRODUCTO'].fillna(cruce['PRODUCTO_SUP']).fillna("N/A")
        cruce['LOTE'] = cruce['LOTE'].fillna(cruce['LOTE_SUP'])
        cruce['SALDO_SAP'] = cruce['SALDO_SAP'].fillna(0).round(2)
        cruce['SALDO_FISICO'] = cruce['SALDO_FISICO'].fillna(0).round(2)
        
        cruce = cruce[~((cruce['SALDO_SAP'] == 0) & (cruce['SALDO_FISICO'] == 0))]
        
        cruce['DIFERENCIA'] = (cruce['SALDO_FISICO'] - cruce['SALDO_SAP']).round(2)
        cruce['ESTADO'] = cruce['DIFERENCIA'].apply(lambda x: "✅ OK" if abs(x) <= 0.05 else "❌ DISCREPANCIA")
        
        cruce['OBSERVACIONES'] = ""
        for idx, row in cruce.iterrows():
            key = f"{row['PISTA']}_{row['LOTE_KEY']}"
            if key in st.session_state.observaciones_memoria:
                cruce.at[idx, 'OBSERVACIONES'] = st.session_state.observaciones_memoria[key]
            elif row['SALDO_SAP'] > 0 and row['SALDO_FISICO'] == 0:
                cruce.at[idx, 'OBSERVACIONES'] = "SUGERIDO: Entrega / Traslado / Pendiente por Facturar"

        st.session_state.cruce_final = cruce[['PISTA', 'ITEM', 'PRODUCTO', 'LOTE_KEY', 'LOTE', 'SALDO_SAP', 'SALDO_FISICO', 'DIFERENCIA', 'ESTADO', 'OBSERVACIONES']].sort_values(by=['PISTA', 'PRODUCTO'])

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Botón en la pantalla principal
    if st.button("🚀 INICIAR ARQUEO ESTRATÉGICO", type="primary", use_container_width=True):
        if not archivo_sap or not archivos_sup or not semana_obj:
            st.error("❌ Faltan suministros. Asegúrese de cargar ambos archivos y escribir la semana.")
        else:
            try:
                with st.spinner("Desplegando analista de inventarios..."):
                    st.session_state.observaciones_memoria = {}
                    
                    sap_file = archivo_sap[0] if isinstance(archivo_sap, list) else archivo_sap
                    nombre_sap = sap_file.name.lower()
                    if nombre_sap.endswith('.xlsx') or nombre_sap.endswith('.xls'):
                        df_sap = pd.read_excel(sap_file)
                    else:
                        try:
                            df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='utf-8')
                        except UnicodeDecodeError:
                            sap_file.seek(0)
                            df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='latin1')

                    df_sap.columns = [quitar_tildes(c) for c in df_sap.columns]
                    
                    c_item = next((c for c in df_sap.columns if "MATERIAL" in c and "DESC" not in c), df_sap.columns[0])
                    c_desc = next((c for c in df_sap.columns if "DESCRIP" in c), df_sap.columns[1])
                    c_pista = next((c for c in df_sap.columns if "ALMACEN" in c or "PISTA" in c), df_sap.columns[2])
                    c_lote = next((c for c in df_sap.columns if "LOTE" in c), df_sap.columns[3])
                    c_saldo = next((c for c in df_sap.columns if "LIBRE" in c or "UTILIZACION" in c), df_sap.columns[4])

                    df_sap_clean = df_sap[[c_item, c_desc, c_pista, c_lote, c_saldo]].copy()
                    df_sap_clean.columns = ['ITEM', 'PRODUCTO', 'PISTA', 'LOTE', 'SALDO_SAP']
                    df_sap_clean['LOTE_KEY'] = df_sap_clean['LOTE'].apply(purificar_lote)
                    df_sap_clean['PISTA'] = df_sap_clean['PISTA'].astype(str).str.strip().str.upper()
                    df_sap_clean['SALDO_SAP'] = pd.to_numeric(df_sap_clean['SALDO_SAP'].astype(str).replace(',', '.'), errors='coerce').fillna(0)
                    
                    st.session_state.df_sap_raw = df_sap_clean 
                    st.session_state.df_sap_grouped = df_sap_clean.groupby(['PISTA', 'LOTE_KEY', 'ITEM', 'PRODUCTO', 'LOTE'], as_index=False)['SALDO_SAP'].sum()

                    lista_sup = []
                    sem_num = str(semana_obj).strip()
                    nombres_pestaña = [sem_num, f"SEM {sem_num}", f"SEM{sem_num}", f"SEMANA {sem_num}"]
                    
                    for file in archivos_sup:
                        dict_dfs = pd.read_excel(file, sheet_name=None, header=None, dtype=str)
                        target = next((n for n in dict_dfs.keys() if str(n).upper().strip() in [p.upper() for p in nombres_pestaña]), None)
                        
                        if target:
                            df_raw = dict_dfs[target]
                            h_idx = -1
                            for i in range(min(30, len(df_raw))):
                                row_v = [quitar_tildes(x) for x in df_raw.iloc[i].values if pd.notna(x)]
                                if any("LOTE" in val for val in row_v) and any("SALDO" in val for val in row_v):
                                    h_idx = i; break
                            if h_idx != -1:
                                df_s = df_raw.iloc[h_idx + 1:].copy()
                                df_s.columns = [f"{quitar_tildes(x)}_{idx}" for idx, x in enumerate(df_raw.iloc[h_idx])]
                                c_p = next((c for c in df_s.columns if "PRODUC" in c or "DESCRI" in c), None)
                                c_a = next((c for c in df_s.columns if "ALMAC" in c or "PISTA" in c), None)
                                c_l = next((c for c in df_s.columns if "LOTE" in c and "SALDO" not in c), None)
                                c_v = next((c for c in df_s.columns if "SALDO" in c and "INIC" not in c), None)
                                if all([c_p, c_a, c_l, c_v]):
                                    df_s_c = df_s[[c_p, c_a, c_l, c_v]].copy()
                                    df_s_c.columns = ['PRODUCTO_SUP', 'PISTA', 'LOTE_SUP', 'SALDO_FISICO']
                                    df_s_c['PISTA'] = df_s_c['PISTA'].astype(str).str.strip().str.upper().replace('NAN', None).ffill().bfill()
                                    df_s_c['LOTE_KEY'] = df_s_c['LOTE_SUP'].apply(purificar_lote)
                                    df_s_c['SALDO_FISICO'] = pd.to_numeric(df_s_c['SALDO_FISICO'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                                    lista_sup.append(df_s_c)

                    if lista_sup:
                        df_sup_total = pd.concat(lista_sup, ignore_index=True)
                        st.session_state.df_sup_grouped = df_sup_total.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        st.session_state.semana_actual = semana_obj
                        generar_cruce()
                        st.session_state.arqueo_procesado = True
                    else:
                        st.error("❌ No se encontraron datos válidos.")

            except Exception as e:
                st.error(f"Error: {e}")
                
    if st.session_state.arqueo_procesado:
        tab1, tab2, tab3 = st.tabs(["⚠️ Discrepancias y Notas", "🛠️ Conciliador Inteligente", "📋 Inventario Completo"])
        
        with tab1:
            st.subheader("Registros con Diferencias (Limpios de 0s)")
            df_err = st.session_state.cruce_final[st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA"].copy()
            
            if df_err.empty:
                st.success("✅ ¡Inventario cuadrado!")
            else:
                edited_df = st.data_editor(
                    df_err.drop(columns=['LOTE_KEY']),
                    use_container_width=True,
                    hide_index=True,
                    disabled=["PISTA", "ITEM", "PRODUCTO", "LOTE", "SALDO_SAP", "SALDO_FISICO", "DIFERENCIA", "ESTADO"],
                    column_config={"OBSERVACIONES": st.column_config.TextColumn("📝 OBSERVACIONES (Editable)", width="large")}
                )
                
                for _, row in edited_df.iterrows():
                    key = f"{row['PISTA']}_{purificar_lote(row['LOTE'])}"
                    st.session_state.observaciones_memoria[key] = row['OBSERVACIONES']
                    idx_m = st.session_state.cruce_final[(st.session_state.cruce_final['PISTA'] == row['PISTA']) & (st.session_state.cruce_final['LOTE_KEY'] == purificar_lote(row['LOTE']))].index
                    if not idx_m.empty:
                        st.session_state.cruce_final.at[idx_m[0], 'OBSERVACIONES'] = row['OBSERVACIONES']

        with tab2:
            st.markdown("### 🛠️ Fusión de Lotes y Nombres Mal Escritos")
            err_fantasmas = st.session_state.cruce_final[(st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA") & (st.session_state.cruce_final['SALDO_SAP'] == 0) & (st.session_state.cruce_final['SALDO_FISICO'] > 0)]
            
            if err_fantasmas.empty:
                st.success("✅ No hay lotes fantasmas pendientes.")
            else:
                opciones = err_fantasmas.apply(lambda x: f"{x['PISTA']} | Prod: {x['PRODUCTO']} | Lote Físico: {x['LOTE']}", axis=1).tolist()
                sel = st.selectbox("1️⃣ Seleccione el error de digitación del supervisor:", opciones)
                
                if sel:
                    idx_s = opciones.index(sel)
                    row_s = err_fantasmas.iloc[idx_s]
                    
                    df_sap_pista = st.session_state.df_sap_raw[st.session_state.df_sap_raw['PISTA'] == row_s['PISTA']]
                    df_exact = df_sap_pista[df_sap_pista['PRODUCTO'] == row_s['PRODUCTO']]
                    
                    if not df_exact.empty:
                        lotes_validos = df_exact.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()
                        lote_ok_str = st.selectbox(f"2️⃣ Lotes Oficiales de SAP para {row_s['PRODUCTO']}:", sorted(lotes_validos))
                    else:
                        st.warning(f"⚠️ El nombre '{row_s['PRODUCTO']}' tiene un error de escritura. Seleccione el producto correcto de esta lista general:")
                        lotes_validos = df_sap_pista.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()
                        lote_ok_str = st.selectbox(f"2️⃣ Arsenal completo de SAP para la pista {row_s['PISTA']}:", sorted(lotes_validos))
                    
                    if st.button("⚡ FUSIONAR Y JUSTIFICAR", type="primary"):
                        prod_sap_oficial = lote_ok_str.split(" | Lote: ")[0].strip()
                        lote_sap_oficial = lote_ok_str.split(" | Lote: ")[1].strip()
                        
                        mask = (st.session_state.df_sup_grouped['PISTA'] == row_s['PISTA']) & (st.session_state.df_sup_grouped['LOTE_KEY'] == row_s['LOTE_KEY'])
                        
                        key_final = f"{row_s['PISTA']}_{purificar_lote(lote_sap_oficial)}"
                        st.session_state.observaciones_memoria[key_final] = f"Corrección: Nombre/Lote Físico ({row_s['PRODUCTO']} - {row_s['LOTE']}) unificado con SAP ({prod_sap_oficial} - {lote_sap_oficial})"
                        
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_SUP'] = lote_sap_oficial
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_KEY'] = purificar_lote(lote_sap_oficial)
                        st.session_state.df_sup_grouped.loc[mask, 'PRODUCTO_SUP'] = prod_sap_oficial
                        
                        st.session_state.df_sup_grouped = st.session_state.df_sup_grouped.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        
                        generar_cruce()
                        st.rerun()

        with tab3:
            st.subheader("Inventario Consolidado (Libre de Ceros)")
            st.dataframe(st.session_state.cruce_final.drop(columns=['LOTE_KEY']).style.map(
                lambda x: 'background-color: #d4edda; color: #155724' if x == "✅ OK" else '', subset=['ESTADO']
            ), use_container_width=True, hide_index=True)

        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            f_df = st.session_state.cruce_final.drop(columns=['LOTE_KEY'])
            f_df[f_df['ESTADO'] == "❌ DISCREPANCIA"].to_excel(writer, index=False, sheet_name='Diferencias')
            f_df.to_excel(writer, index=False, sheet_name='Total')
            
            for sheetname in writer.sheets:
                worksheet = writer.sheets[sheetname]
                worksheet.auto_filter.ref = worksheet.dimensions 
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 40)
                    worksheet.column_dimensions[column].width = adjusted_width

        st.download_button("📥 Descargar Reporte Ejecutivo", buffer.getvalue(), f"Arqueo_Ejecutivo_Semana_{st.session_state.semana_actual}.xlsx")

# =====================================================================
# 📊 8. REPORTE TÁCTICO DE HECTÁREAS FUMIGADAS
# =====================================================================
elif menu == "📊 8. Reporte Hectáreas (Pistas)":
    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
    try:
        with st.spinner("🛰️ Escaneando la Bóveda Maestra (TABLA 1)..."):
            if "gcp_credentials" in st.secrets:
                gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
            else:
                gc = gspread.service_account(filename='credenciales.json')
            
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            hoja_maestra = boveda.worksheet("TABLA 1")
            datos_brutos = hoja_maestra.get_all_values()
            
        if len(datos_brutos) > 5:
            columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "HA_NETAS", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "H_PROPORCIONAL", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_TOTAL_AVION", "TARIFA_HA", "RECARGO_HA", "SUBTOTAL", "COSTO_HORA", "PISTA"]
            
            filas_limpias = [r + [""]*(24 - len(r)) for r in datos_brutos[5:]]
            df_rep = pd.DataFrame([r[:24] for r in filas_limpias], columns=columnas)
            
            df_rep['HA_NETAS'] = df_rep['HA_NETAS'].apply(extraer_numero)
            df_rep['H_PROPORCIONAL'] = df_rep['H_PROPORCIONAL'].apply(extraer_numero)
            df_rep['SEMANA'] = df_rep['SEMANA'].astype(str).str.strip()
            df_rep['PISTA'] = df_rep['PISTA'].astype(str).str.strip().str.upper()
            
            df_rep = df_rep[(df_rep['PISTA'] != "") & (df_rep['HA_NETAS'] > 0)]
            
            meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
            
            def extraer_mes_año(fecha_str):
                dt = procesar_fecha_pesada(fecha_str)
                if dt: return meses_nom.get(dt.month, "00-Desc"), str(dt.year)
                return "00-Desc", "00-Desc"
            
            df_rep[['MES', 'AÑO']] = df_rep['FECHA'].apply(lambda x: pd.Series(extraer_mes_año(x)))
            df_rep = df_rep[df_rep['AÑO'] != "00-Desc"]
            
            # --- PANEL DE CONTROL ---
            st.markdown("### 🎛️ Centro de Comando y Filtros")
            c1, c2, c3 = st.columns([2, 1, 1])
            
            vista_seleccionada = c1.radio(
                "👁️ Seleccione la Vista del Radar:", 
                ["📊 Resumen Gerencial (Hectáreas)", "📅 Mapa Semanal (Detalle)"], 
                horizontal=True
            )
            
            pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
            años_disp = sorted(df_rep['AÑO'].unique().tolist(), reverse=True)
            
            año_sel = c2.selectbox("📅 Año Fiscal", años_disp if años_disp else [str(datetime.now().year)])
            pista_sel = c3.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp)
            
            # ⚡ BOTÓN SECRETO PARA MOSTRAR HORAS
            mostrar_horas = False
            if vista_seleccionada == "📊 Resumen Gerencial (Hectáreas)":
                mostrar_horas = st.checkbox("⏱️ Mostrar también el Rendimiento (Horas de Vuelo)")

            df_filt = df_rep[df_rep['AÑO'] == año_sel]
            if pista_sel != "TODAS":
                df_filt = df_filt[df_filt['PISTA'] == pista_sel]
            
            if df_filt.empty:
                st.warning("⚠️ No hay operaciones registradas para estos parámetros.")
            else:
                st.markdown("---")
                
                # =========================================================
                # VISTA 1: RESUMEN GERENCIAL
                # =========================================================
                if vista_seleccionada == "📊 Resumen Gerencial (Hectáreas)":
                    st.markdown(f"#### 📑 Consolidado Gerencial - {año_sel}")
                    
                    df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(
                        REND_HR=('H_PROPORCIONAL', 'sum'),
                        AREA_FUMIG=('HA_NETAS', 'sum')
                    ).reset_index()
                    
                    tabla_final = []
                    total_hr_gral = 0
                    total_ha_gral = 0
                    
                    for pista in sorted(df_gerencia['PISTA'].unique()):
                        datos_pista = df_gerencia[df_gerencia['PISTA'] == pista].sort_values(by='MES')
                        sum_hr = datos_pista['REND_HR'].sum()
                        sum_ha = datos_pista['AREA_FUMIG'].sum()
                        
                        fila_sub = {'NIVEL': f"➖ {pista}", 'MES': ''}
                        if mostrar_horas: fila_sub['REND (hr)'] = sum_hr
                        fila_sub['ÁREA FUMIG (ha)'] = sum_ha
                        tabla_final.append(fila_sub)
                        
                        for _, row in datos_pista.iterrows():
                            mes_limpio = row['MES'].split('-')[1] if '-' in row['MES'] else row['MES']
                            fila_mes = {'NIVEL': '', 'MES': mes_limpio}
                            if mostrar_horas: fila_mes['REND (hr)'] = row['REND_HR']
                            fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                            tabla_final.append(fila_mes)
                            
                        total_hr_gral += sum_hr
                        total_ha_gral += sum_ha
                        
                    fila_tot = {'NIVEL': 'TOTAL GENERAL', 'MES': ''}
                    if mostrar_horas: fila_tot['REND (hr)'] = total_hr_gral
                    fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                    tabla_final.append(fila_tot)
                    
                    df_visual = pd.DataFrame(tabla_final)
                    
                    def estilizar_filas(row):
                        if "➖" in row['NIVEL'] or "TOTAL" in row['NIVEL']:
                            return ['background-color: #e2e6ea; font-weight: bold;'] * len(row)
                        return [''] * len(row)
                    
                    formato_columnas = {'ÁREA FUMIG (ha)': "{:.2f}"}
                    if mostrar_horas: formato_columnas['REND (hr)'] = "{:.2f}"
                    
                    st.dataframe(
                        df_visual.style.apply(estilizar_filas, axis=1).format(formato_columnas),
                        use_container_width=True,
                        hide_index=True
                    )

                # =========================================================
                # VISTA 2: MAPA SEMANAL (EL CALOR)
                # =========================================================
                else:
                    matriz = pd.pivot_table(df_filt, values='HA_NETAS', index='MES', columns='SEMANA', aggfunc='sum', fill_value=0)
                    matriz = matriz.sort_index()
                    cols_ordenadas = sorted(matriz.columns, key=lambda x: int(x) if str(x).isdigit() else 999)
                    matriz = matriz[cols_ordenadas]
                    
                    matriz.index = [m.split('-')[1] if '-' in m else m for m in matriz.index]
                    matriz['TOTAL MES'] = matriz.sum(axis=1)
                    matriz.loc['TOTAL ANUAL'] = matriz.sum(axis=0)
                    
                    st.markdown(f"#### 🚜 Rendimiento Semana a Semana: **{pista_sel}**")
                    if HAS_MATPLOTLIB:
                        st.dataframe(matriz.style.format("{:.2f}").background_gradient(cmap="YlGn", axis=None), use_container_width=True)
                    else:
                        st.dataframe(matriz.style.format("{:.2f}"), use_container_width=True)
                    
                    st.markdown("---")
                    df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
                    if not df_grafico.empty:
                        fig = px.bar(
                            df_grafico, x='index', y='TOTAL MES', text='TOTAL MES',
                            labels={'TOTAL MES': 'Hectáreas Fumigadas', 'index': 'Mes de Operación'},
                            color='TOTAL MES', color_continuous_scale='Greens'
                        )
                        fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                        fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', showlegend=False, xaxis_title="Mes")
                        st.plotly_chart(fig, use_container_width=True)

                # --- BOTÓN DE EXPORTACIÓN: VERSIÓN MASTER VISIBILIDAD ---
                st.markdown("---")
                buffer_rep = io.BytesIO()
                with pd.ExcelWriter(buffer_rep, engine='openpyxl') as writer:
                    nombre_hoja = 'Resumen_Gerencial' if "Gerencial" in vista_seleccionada else 'Reporte_Semanal'
                    
                    # 1. Inyectamos la estructura base
                    if "Gerencial" in vista_seleccionada:
                        df_visual.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    else:
                        matriz.to_excel(writer, sheet_name=nombre_hoja)
                        
                    workbook = writer.book
                    worksheet = writer.sheets[nombre_hoja]
                    
                    # Estética Dashboard: Sin líneas de cuadrícula y filas más altas
                    worksheet.sheet_view.showGridLines = False
                    worksheet.row_dimensions[1].height = 30 # Encabezado alto
                    
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.chart import BarChart, Reference
                    from openpyxl.chart.label import DataLabelList
                    from openpyxl.utils import get_column_letter

                    # --- ESTILOS DE GALA ---
                    borde_pro = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), 
                                       top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                    fondo_navy = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                    fuente_blanca = Font(color="FFFFFF", bold=True, size=11)
                    fondo_meses = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
                    fondo_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    fondo_total = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")

                    max_row = worksheet.max_row
                    max_col = worksheet.max_column

                    # --- 🔄 INYECCIÓN DE FÓRMULAS VIVAS (SUMAS AUTOMÁTICAS) ---
                    if "Gerencial" in vista_seleccionada:
                        rango_total_ha = []
                        col_ha_letra = "C" if not mostrar_horas else "D"
                        col_ha_idx = 3 if not mostrar_horas else 4
                        
                        for i in range(2, max_row + 1):
                            nivel = str(worksheet.cell(row=i, column=1).value or "").strip()
                            
                            if "➖" in nivel:
                                inicio = i + 1
                                fin = i + 1
                                for j in range(i + 1, max_row + 1):
                                    val_j = str(worksheet.cell(row=j, column=1).value or "").strip()
                                    if val_j == "" or val_j == "None": fin = j
                                    else: break
                                # Fórmula de Suma para la Pista
                                worksheet.cell(row=i, column=col_ha_idx).value = f"=SUM({col_ha_letra}{inicio}:{col_ha_letra}{fin})"
                                rango_total_ha.append(f"{col_ha_letra}{i}")
                                
                            elif "TOTAL GENERAL" in nivel:
                                if rango_total_ha:
                                    worksheet.cell(row=i, column=col_ha_idx).value = f"=SUM({','.join(rango_total_ha)})"

                    # --- APLICAR DISEÑO PROFESIONAL ---
                    for row in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                        for cell in row:
                            cell.border = borde_pro
                            # Formato numérico
                            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.number_format = '#,##0.00'
                            
                            # Estilo de Fila 1 (Encabezados)
                            if cell.row == 1:
                                cell.fill = fondo_navy; cell.font = fuente_blanca
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            else:
                                cell.alignment = Alignment(vertical='center', indent=1)
                                
                            # Colores de jerarquía
                            if "Gerencial" in vista_seleccionada and cell.row > 1:
                                nivel_v = str(worksheet.cell(row=cell.row, column=1).value or "").strip()
                                if "➖" in nivel_v:
                                    cell.fill = fondo_sub; cell.font = Font(bold=True)
                                elif "TOTAL GENERAL" in nivel_v:
                                    cell.fill = fondo_total; cell.font = Font(bold=True, color="FFFFFF")
                                elif nivel_v == "" or nivel_v == "None":
                                    cell.fill = fondo_meses

                    # --- 📈 GRÁFICO VIVO (Ubicado a la derecha para no tapar la tabla) ---
                    chart = BarChart()
                    chart.type = "col"; chart.style = 10
                    chart.title = "Rendimiento Operativo (Ha)"; chart.y_axis.title = "Hectáreas"
                    chart.legend = None
                    chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True # Números sobre las barras
                    chart.height = 14; chart.width = 24
                    
                    if "Gerencial" in vista_seleccionada:
                        # Minitabla invisible para el gráfico en AA y AB
                        worksheet.cell(row=1, column=27).value = "Mes"
                        worksheet.cell(row=1, column=28).value = "Ha"
                        
                        meses_para_grafico = [m for m in df_visual['MES'] if str(m).strip() not in ["", "None"]]
                        row_g = 2
                        for m in meses_para_grafico:
                            worksheet.cell(row=row_g, column=27).value = m
                            # Buscamos la fila de este mes en la tabla principal para linkear la fórmula
                            fila_origen = 2
                            for r_b in range(2, max_row):
                                if str(worksheet.cell(row=r_b, column=2).value) == m:
                                    fila_origen = r_b; break
                            worksheet.cell(row=row_g, column=28).value = f"={col_ha_letra}{fila_origen}"
                            row_g += 1
                        
                        data = Reference(worksheet, min_col=28, min_row=1, max_row=row_g-1)
                        cats = Reference(worksheet, min_col=27, min_row=2, max_row=row_g-1)
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        
                        # Invisible: Fuente Blanca
                        for r_inv in range(1, row_g):
                            worksheet.cell(row=r_inv, column=27).font = Font(color="FFFFFF")
                            worksheet.cell(row=r_inv, column=28).font = Font(color="FFFFFF")
                        
                        # POSICIÓN: Empezar en la Columna H para dar visibilidad total a la tabla
                        worksheet.add_chart(chart, "H2")
                    else:
                        data = Reference(worksheet, min_col=max_col, min_row=1, max_row=max_row-1)
                        cats = Reference(worksheet, min_col=1, min_row=2, max_row=max_row-1)
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        worksheet.add_chart(chart, f"{get_column_letter(max_col + 2)}2")
                    
                    # Ajuste de ancho de columnas
                    for col_idx in range(1, max_col + 1):
                        worksheet.column_dimensions[get_column_letter(col_idx)].width = 22
                    worksheet.freeze_panes = "A2"

                st.download_button(
                    label="💾 DESCARGAR REPORTE GERENCIAL TOP",
                    data=buffer_rep.getvalue(),
                    file_name=f"Reporte_Gerencial_{año_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )                
    except Exception as e:
        st.error(f"🚨 Falla en el sistema de radares: {e}")



# =====================================================================
# 📈 9. DASHBOARD TÁCTICO (FUSIÓN EXCEL + STREAMLIT)
# =====================================================================
elif menu == "📈 9. Dashboard Táctico":
    st.markdown("<h1 class='titulo-principal'>Centro de Comando: Rendimiento y Finanzas</h1>", unsafe_allow_html=True)
    
    import plotly.graph_objects as go
    import plotly.express as px

    with st.spinner("📡 Conectando con la Bóveda de Datos (TABLA 1)..."):
        try:
            # 1. CONEXIÓN A LA BASE DE DATOS
            if "gcp_credentials" in st.secrets:
                gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
            else:
                gc = gspread.service_account(filename='credenciales.json')
            
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            hoja_maestra = boveda.worksheet("TABLA 1")
            datos_brutos = hoja_maestra.get_all_values()
            
            if len(datos_brutos) > 5:
                columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA", "INC_2026", "LIMITE", "ALERTA", "VAR_PCT", "COSTO_TOTAL", "PAGO_AVION"]
                
                filas_limpias = [r + [""]*(len(columnas) - len(r)) for r in datos_brutos[5:]]
                df_dash = pd.DataFrame([r[:len(columnas)] for r in filas_limpias], columns=columnas)
                
                cols_numericas = ['AREA_FUMIG', 'REND_HR', 'COSTO_HA', 'DOMINICAL_HA', 'VALOR_FACTURAR', 'LIMITE', 'COSTO_TOTAL', 'COSTO_AVION']
                for col in cols_numericas:
                    df_dash[col] = df_dash[col].apply(extraer_numero)
                
                df_dash['FECHA_DT'] = df_dash['FECHA'].apply(procesar_fecha_pesada)
                df_dash = df_dash.dropna(subset=['FECHA_DT'])
                
                # 🎯 NUEVA INTELIGENCIA TEMPORAL
                df_dash['AÑO'] = df_dash['FECHA_DT'].dt.year
                df_dash['TRIMESTRE'] = df_dash['FECHA_DT'].dt.quarter
                df_dash['MES_NUM'] = df_dash['FECHA_DT'].dt.month
                meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
                df_dash['MES_NOMBRE'] = df_dash['MES_NUM'].map(meses_dict)
                df_dash['MES_ORDEN'] = df_dash['AÑO'].astype(str) + "-" + df_dash['MES_NUM'].astype(str).str.zfill(2) + " (" + df_dash['MES_NOMBRE'] + ")"
                
                df_dash = df_dash[df_dash['AREA_FUMIG'] > 0] 

                # --- 🎛️ FILTROS TÁCTICOS AVANZADOS ---
                st.markdown("### 🎛️ Filtros de Operación y Tiempo")
                
                # Fila 1: Filtros de Tiempo
                t1, t2 = st.columns(2)
                # 🎯 AÑADIMOS "TODOS" AL INICIO DE LA LISTA DE AÑOS
                años_disp = ["TODOS"] + sorted(df_dash['AÑO'].astype(int).unique().tolist(), reverse=True)
                año_sel = t1.selectbox("📅 AÑO FISCAL", años_disp, index=0)
                
                trimestres = {"TODOS": 0, "Q1 (Ene-Mar)": 1, "Q2 (Abr-Jun)": 2, "Q3 (Jul-Sep)": 3, "Q4 (Oct-Dic)": 4}
                trim_sel = t2.selectbox("📊 TRIMESTRE", list(trimestres.keys()))

                # Fila 2: Filtros Operativos
                f1, f2, f3 = st.columns(3)
                fincas_disp = ["TODAS"] + sorted(df_dash['FINCA'].astype(str).unique().tolist())
                pilotos_disp = ["TODOS"] + sorted(df_dash['PILOTO'].astype(str).unique().tolist())
                hks_disp = ["TODAS"] + sorted(df_dash['HK'].astype(str).unique().tolist())
                
                finca_filtro = f1.selectbox("📍 FINCA", fincas_disp)
                piloto_filtro = f2.selectbox("👨‍✈️ PILOTO", pilotos_disp)
                hk_filtro = f3.selectbox("✈️ MATRÍCULA (HK)", hks_disp)

                # 🎯 APLICAR FILTROS (Lógica actualizada para aceptar "TODOS" en Años)
                df_filtrado = df_dash.copy()
                if año_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['AÑO'] == año_sel]
                if trimestres[trim_sel] != 0: df_filtrado = df_filtrado[df_filtrado['TRIMESTRE'] == trimestres[trim_sel]]
                if finca_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['FINCA'] == finca_filtro]
                if piloto_filtro != "TODOS": df_filtrado = df_filtrado[df_filtrado['PILOTO'] == piloto_filtro]
                if hk_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['HK'] == hk_filtro]

                # --- 🏆 TARJETAS DE MANDO (KPIs) ---
                total_area = df_filtrado['AREA_FUMIG'].max() if not df_filtrado.empty else 0
                
                # 🎯 CORRECCIÓN: Sumar directamente COSTO_TOTAL (La columna de los Millones)
                total_facturacion = float(df_filtrado['COSTO_TOTAL'].sum())
                
                total_dominical = float(df_filtrado['DOMINICAL_HA'].sum())
                
                st.markdown("<br>", unsafe_allow_html=True)
                k1, k2, k3 = st.columns(3)
                
                estilo_kpi = "background-color: #D9E1F2; border: 2px solid #2F75B5; border-radius: 10px; padding: 15px; text-align: center;"
                k1.markdown(f"<div style='{estilo_kpi}'><h4 style='color:#0d1b2a; margin:0;'>🚜 ÁREA FINCA (Ha)</h4><h2 style='color:#2F75B5; margin:0;'>{total_area:,.2f}</h2></div>", unsafe_allow_html=True)
                k2.markdown(f"<div style='{estilo_kpi}'><h4 style='color:#0d1b2a; margin:0;'>💰 FACTURACIÓN TOTAL</h4><h2 style='color:#2F75B5; margin:0;'>$ {total_facturacion:,.0f}</h2></div>", unsafe_allow_html=True)
                k3.markdown(f"<div style='{estilo_kpi}'><h4 style='color:#0d1b2a; margin:0;'>⚠️ DOMINICALES TOTAL</h4><h2 style='color:#2F75B5; margin:0;'>$ {total_dominical:,.0f}</h2></div>", unsafe_allow_html=True)

                st.markdown("<hr>", unsafe_allow_html=True)
                
                if df_filtrado.empty:
                    st.warning(f"⚠️ El Escuadrón no registró operaciones con los filtros actuales.")
                else:
                    g1, g2 = st.columns(2)

                    # --- GRÁFICO 1: ÁREA ASPERJADA (Agrupada por MES) ---
                    with g1:
                        st.markdown(f"<h4 style='text-align:center;'>🚜 ÁREA ASPERJADA POR MES</h4>", unsafe_allow_html=True)
                        df_area = df_filtrado.groupby('MES_ORDEN')['AREA_FUMIG'].sum().reset_index()
                        df_area = df_area.sort_values(by='MES_ORDEN')
                        
                        fig1 = px.bar(df_area, x='MES_ORDEN', y='AREA_FUMIG', text='AREA_FUMIG', color_discrete_sequence=['#548235'])
                        fig1.update_traces(texttemplate='%{text:.1f}', textposition='outside', textfont_size=14)
                        fig1.update_layout(xaxis_title="Mes Operativo", yaxis_title="Hectáreas", plot_bgcolor='rgba(0,0,0,0)', uniformtext_minsize=12)
                        st.plotly_chart(fig1, use_container_width=True)

                    # --- GRÁFICO 2: FACTURACIÓN/ha vs LÍMITE (Corregido y Optimizado Visualmente) ---
                    with g2:
                        st.markdown(f"<h4 style='text-align:center;'>⚖️ FACTURACIÓN/ha vs LÍMITE</h4>", unsafe_allow_html=True)
                        
                        df_costo = df_filtrado.groupby(['MES_ORDEN', 'COCTEL']).agg({
                            'VALOR_FACTURAR': 'mean', 
                            'LIMITE': 'max'
                        }).reset_index()
                        
                        limite_real = df_filtrado[df_filtrado['LIMITE'] > 0]['LIMITE'].max()
                        if pd.isna(limite_real) or limite_real == 0: 
                            limite_real = 200000 
                            
                        df_costo['LIMITE'] = df_costo['LIMITE'].apply(lambda x: limite_real if x == 0 else x)
                        
                        def acortar_fecha(txt):
                            try: return txt.split('(')[1].replace(')','') + " '" + txt[2:4]
                            except: return txt
                            
                        df_costo['FECHA_CORTA'] = df_costo['MES_ORDEN'].apply(acortar_fecha)
                        
                        # ✂️ TRUCO MÁS AGRESIVO: Máximo 10 letras para el cóctel en el eje
                        df_costo['COCTEL_CORTO'] = df_costo['COCTEL'].apply(lambda x: str(x)[:10] + '..' if len(str(x)) > 10 else str(x))
                        df_costo['ETIQUETA'] = df_costo['COCTEL_CORTO'] + "<br>(" + df_costo['FECHA_CORTA'] + ")"

                        fig2 = go.Figure()
                        
                        fig2.add_trace(go.Bar(
                            x=df_costo['ETIQUETA'], 
                            y=df_costo['VALOR_FACTURAR'], 
                            name="Facturación/ha",
                            marker_color='#548235', 
                            text=df_costo['VALOR_FACTURAR'], 
                            texttemplate='$ %{text:,.0f}', 
                            textposition='outside', 
                            textfont=dict(size=11), # Letra del valor un poco más pequeña
                            hovertext=df_costo['COCTEL'], 
                            hovertemplate='<b>Cóctel:</b> %{hovertext}<br><b>Facturación:</b> $ %{y:,.0f} COP<extra></extra>'
                        ))
                        
                        fig2.add_trace(go.Scatter(
                            x=df_costo['ETIQUETA'], 
                            y=df_costo['LIMITE'], 
                            name="Límite Finca",
                            mode='lines+markers', 
                            line=dict(color='red', width=3), 
                            marker=dict(size=8),
                            hovertemplate='<b>Límite Fijo:</b> $ %{y:,.0f} COP<extra></extra>'
                        ))
                        
                        fig2.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', 
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            yaxis=dict(title="Valor ($ COP / Ha)", rangemode='tozero'),
                            margin=dict(b=100) # 🎯 MÁS ESPACIO ABAJO para el texto vertical
                        )
                        # 🎯 TEXTO COMPLETAMENTE VERTICAL (-90 GRADOS) Y TAMAÑO REDUCIDO
                        fig2.update_xaxes(tickangle=-90, tickfont=dict(size=10)) 
                        st.plotly_chart(fig2, use_container_width=True)
                        
                    g3, g4 = st.columns(2)

                    # --- GRÁFICO 3: RENDIMIENTO/Hora FINCA (Estilo Excel) ---
                    with g3:
                        titulo_finca = f" {finca_filtro}" if finca_filtro != "TODAS" else ""
                        st.markdown(f"<h4 style='text-align:center;'>⏱️ RENDIMIENTO/Hora FINCA{titulo_finca}</h4>", unsafe_allow_html=True)
                        
                        # 1. Agrupamos por Máquina (HK) y Semana, igual que en Excel
                        df_rend = df_filtrado.groupby(['HK', 'SEMANA'])['REND_HR'].sum().reset_index()
                        
                        # 2. TRUCO VITAL: Convertimos HK y Semana a texto puro para que Python no los sume ni los aplaste
                        df_rend['HK'] = df_rend['HK'].astype(str).str.replace(".0", "", regex=False)
                        df_rend['SEMANA'] = df_rend['SEMANA'].astype(str).str.replace(".0", "", regex=False)
                        
                        # 3. Creamos la jerarquía visual para el eje Y (Ej: "4014 | Sem 2")
                        df_rend['EJE_Y'] = df_rend['HK'] + " | Sem " + df_rend['SEMANA']
                        
                        # Ordenamos para que las máquinas y las semanas queden agrupadas y en orden
                        df_rend = df_rend.sort_values(by=['HK', 'SEMANA'], ascending=[True, False])
                        
                        # 4. Generamos el gráfico
                        fig3 = px.bar(df_rend, y='EJE_Y', x='REND_HR', orientation='h', text='REND_HR',
                                      color_discrete_sequence=['#548235'])
                        
                        # Formato de los números sobre las barras
                        fig3.update_traces(texttemplate='%{text:.2f}', textposition='outside', textfont_size=14)
                        
                        # Ajuste de fondo y títulos
                        fig3.update_layout(yaxis_title="Matrícula (HK) | Semana", xaxis_title="Rendimiento (Horas)", plot_bgcolor='rgba(0,0,0,0)')
                        
                        # BLINDAJE: Forzar el eje Y para que trate las etiquetas como categorías
                        fig3.update_yaxes(type='category')
                        
                        st.plotly_chart(fig3, use_container_width=True)
                    # --- GRÁFICO 4: FACTURACIÓN MENSUAL (Corregido) ---
                    with g4:
                        st.markdown(f"<h4 style='text-align:center;'>💵 FACTURACIÓN MENSUAL</h4>", unsafe_allow_html=True)
                        
                        # 🎯 CORRECCIÓN: Agrupamos y sumamos COSTO_TOTAL (Millones) en lugar del valor unitario
                        df_mes = df_filtrado.groupby('MES_ORDEN')['COSTO_TOTAL'].sum().reset_index().sort_values(by='MES_ORDEN')
                        
                        fig4 = px.bar(df_mes, x='MES_ORDEN', y='COSTO_TOTAL', text='COSTO_TOTAL',
                                      color_discrete_sequence=['#548235'])
                        
                        # Mostramos el valor en formato moneda
                        fig4.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont_size=14)
                        fig4.update_layout(xaxis_title="Mes Operativo", yaxis_title="Total Facturado ($)", plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig4, use_container_width=True)
        except Exception as e:
            st.error(f"🚨 Falla en los motores del Dashboard: {e}")


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
        
    # 2. 🎯 ESTANDARIZADOR BLINDADO: Prioridad Máxima
    def estandarizar_base(df):
        renombres = {}
        for col in df.columns:
            # Reemplazamos saltos de línea invisibles por espacios
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
            # 🎯 BUSCADOR LÁSER EXCLUSIVO: Busca la palabra FUMIG
            elif not area_ok and ('FUMIG' in col_u or 'AREA' in col_u or col_u == 'HAS'):
                renombres[col] = 'AREA_MAESTRA'
                area_ok = True
                
        df.rename(columns=renombres, inplace=True)
        return df
        
    # 3. 🎯 TRADUCTOR FINANCIERO SUPERIOR (A prueba de letras)
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
    with st.spinner("📡 Sincronizando Bóveda Maestra y Archivo Histórico..."):
        try:
            if "gcp_credentials" in st.secrets: gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
            else: gc = gspread.service_account(filename='credenciales.json')
                
            # 🟢 CANAL A: Datos Vivos
            boveda_actual = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            datos_brutos_act = boveda_actual.worksheet("TABLA 1").get_all_values()
            
            if len(datos_brutos_act) > 5:
                df_vivos = pd.DataFrame(datos_brutos_act[5:], columns=datos_brutos_act[4])
                df_vivos = estandarizar_base(limpiar_encabezados(df_vivos))
                df_vivos['ORIGEN_BI'] = 'ACTUAL'
            else: df_vivos = pd.DataFrame()

            # 🔵 CANAL B: Datos Históricos
            boveda_hist = gc.open_by_url("https://docs.google.com/spreadsheets/d/16OZdiWwW7nLHyZBEnhiKlDTDttR7Tjhn37O9zm6wJOk/edit")
            try: hoja_hist = boveda_hist.worksheet("Datos")
            except: hoja_hist = boveda_hist.get_worksheet(0)
                
            datos_brutos_hist = hoja_hist.get_all_values()
            if len(datos_brutos_hist) > 0:
                df_historico = pd.DataFrame(datos_brutos_hist[1:], columns=datos_brutos_hist[0])
                df_historico = estandarizar_base(limpiar_encabezados(df_historico))
                df_historico['ORIGEN_BI'] = 'HISTORICO'
            else: df_historico = pd.DataFrame()

            # 🤝 FUSIÓN DEFINITIVA
            if not df_vivos.empty and not df_historico.empty:
                columnas_comunes = list(set(df_vivos.columns).intersection(set(df_historico.columns)))
                if 'ORIGEN_BI' in columnas_comunes: columnas_comunes.remove('ORIGEN_BI')
                
                # Validamos que nuestras columnas clave hayan sobrevivido
                if 'COSTO_MAESTRO' in columnas_comunes and 'FINCA_MAESTRA' in columnas_comunes:
                    df_vivos_trim = df_vivos[columnas_comunes + ['ORIGEN_BI']].copy()
                    df_historico_trim = df_historico[columnas_comunes + ['ORIGEN_BI']].copy()
                    
                    super_base_bi = pd.concat([df_historico_trim, df_vivos_trim], ignore_index=True)
                    super_base_bi['FINCA_MAESTRA'] = super_base_bi['FINCA_MAESTRA'].astype(str).str.strip().str.upper()

                    # =====================================================================
                    # =====================================================================
                    # =====================================================================
                    # --- ⚙️ FASE 3: MOTOR DE TIEMPO Y FILTROS TÁCTICOS ---
                    # =====================================================================
                    st.markdown("---")
                    st.markdown("### 🎛️ Centro de Mando: Parámetros de Análisis")
                    # ... (código intermedio) ...
                    
                    if 'FECHA_MAESTRA' in super_base_bi.columns:
                        super_base_bi['FECHA_DT'] = super_base_bi['FECHA_MAESTRA'].apply(procesar_fecha_pesada)
                        super_base_bi = super_base_bi.dropna(subset=['FECHA_DT'])
                        
                        # Extracción de inteligencia temporal
                        super_base_bi['AÑO'] = super_base_bi['FECHA_DT'].dt.year.astype(int)
                        super_base_bi['MES'] = super_base_bi['FECHA_DT'].dt.month.astype(int)
                        super_base_bi['TRIMESTRE'] = super_base_bi['FECHA_DT'].dt.quarter.astype(int)
                        
                        fincas_disp = ["TODAS"] + sorted(super_base_bi['FINCA_MAESTRA'].dropna().unique().tolist())
                        años_disp = sorted(super_base_bi['AÑO'].unique().tolist(), reverse=True)
                        
                        # Escáner de Escuadrón (Dron vs Avión)
                        col_modelo = 'MODELO' if 'MODELO' in super_base_bi.columns else None
                        if col_modelo:
                            super_base_bi[col_modelo] = super_base_bi[col_modelo].astype(str).str.strip().str.upper()
                            modelos_disp = ["TODOS"] + sorted(super_base_bi[col_modelo].unique().tolist())
                        else:
                            modelos_disp = ["TODOS"]
                        
                        # FILA 1: Objetivos Físicos
                        f1, f2 = st.columns(2)
                        finca_sel = f1.selectbox("📍 Objetivo Geográfico (Finca)", fincas_disp)
                        modelo_sel = f2.selectbox("🚁 Escuadrón (Modelo/Tipo)", modelos_disp)
                        
                        # FILA 2: Lupa Temporal (El viaje en el tiempo granular)
                        t1, t2, t3, t4 = st.columns(4)
                        idx_base = 1 if len(años_disp) > 1 else 0
                        año_base = t1.selectbox("📅 Año Base (Referencia)", años_disp, index=idx_base)
                        año_comp = t2.selectbox("📆 Año Actual (Evaluar)", años_disp, index=0)
                        
                        tipo_periodo = t3.selectbox("⏱️ Lupa Temporal", ["AÑO COMPLETO", "POR TRIMESTRE", "POR MES"])
                        
                        meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
                        
                        if tipo_periodo == "POR TRIMESTRE":
                            periodo_sel = t4.selectbox("📊 Seleccione Trimestre", [1, 2, 3, 4], format_func=lambda x: f"Q{x}")
                        elif tipo_periodo == "POR MES":
                            periodo_sel = t4.selectbox("📅 Seleccione Mes", list(meses_dict.keys()), format_func=lambda x: meses_dict[x])
                        else:
                            t4.markdown("<br><span style='color:gray;'>Visión Anual Activada</span>", unsafe_allow_html=True)
                            periodo_sel = "TODOS"

                        # 🎯 APLICACIÓN DE FILTROS EN CASCADA
                        df_finca = super_base_bi.copy()
                        if finca_sel != "TODAS": df_finca = df_finca[df_finca['FINCA_MAESTRA'] == finca_sel]
                        if col_modelo and modelo_sel != "TODOS": df_finca = df_finca[df_finca[col_modelo] == modelo_sel]
                            
                        df_finca['COSTO_NUM'] = df_finca['COSTO_MAESTRO'].apply(convertir_pesos)

                        # División del Espacio-Tiempo
                        df_periodo_a = df_finca[df_finca['AÑO'] == año_base]
                        df_periodo_b = df_finca[df_finca['AÑO'] == año_comp]
                        
                        # 🎯 CORTE QUIRÚRGICO (Mes a Mes o Trimestre a Trimestre)
                        if tipo_periodo == "POR TRIMESTRE":
                            df_periodo_a = df_periodo_a[df_periodo_a['TRIMESTRE'] == periodo_sel]
                            df_periodo_b = df_periodo_b[df_periodo_b['TRIMESTRE'] == periodo_sel]
                            etiq_periodo = f"Q{periodo_sel}"
                        elif tipo_periodo == "POR MES":
                            df_periodo_a = df_periodo_a[df_periodo_a['MES'] == periodo_sel]
                            df_periodo_b = df_periodo_b[df_periodo_b['MES'] == periodo_sel]
                            etiq_periodo = meses_dict[periodo_sel]
                        else:
                            etiq_periodo = "Total"

                        # Cálculos de impacto
                        costo_a = df_periodo_a['COSTO_NUM'].mean() if not df_periodo_a.empty else 0
                        costo_b = df_periodo_b['COSTO_NUM'].mean() if not df_periodo_b.empty else 0
                        
                        delta_pct = ((costo_b - costo_a) / costo_a * 100) if costo_a > 0 else 0
                        
                        # 7. Artillería Visual: Tarjetas de Impacto
                        st.markdown("### 📊 Auditoría de Costos: Impacto General por Hectárea")
                        
                        k1, k2, k3 = st.columns(3)
                        k1.metric(label=f"Costo Promedio Ha ({año_base})", value=f"$ {costo_a:,.0f}")
                        k2.metric(label=f"Costo Promedio Ha ({año_comp})", value=f"$ {costo_b:,.0f}")
                        k3.metric(label="Variación Total (%)", value=f"{delta_pct:+.2f} %", delta=f"{delta_pct:+.2f}%", delta_color="inverse")
                        
                        # 8. Sistema de Alerta Temprana (Tono Ejecutivo)
                        st.markdown("<br>", unsafe_allow_html=True)
                        if delta_pct > 10:
                            st.error(f"⚠️ **ALERTA ROJA:** El costo operativo en {finca_sel} presenta una desviación del **{delta_pct:.1f}%**. Se requiere análisis de causa raíz.")
                        elif delta_pct < 0:
                            st.success(f"✅ **RENDIMIENTO ÓPTIMO:** El costo operativo se redujo. Excelente gestión logística.")
                        else:
                            st.info(f"⚖️ **ESTABILIDAD:** Los costos se mantienen dentro de los márgenes normales de variación.")
                            # =====================================================================
                        # --- ⏱️ ANEXO TÁCTICO: FRECUENCIA OPERATIVA (CICLOS E INTERVALOS) ---
                        # =====================================================================
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("#### ⏱️ Análisis de Frecuencia: Ciclos Reales e Intervalo")
                        
                        # 🎯 Motor de cálculo de CICLOS (Blindado contra Drones y clima)
                        def calcular_frecuencia(df):
                            if df.empty or 'FECHA_DT' not in df.columns: return 0, 0
                            fechas = sorted(df['FECHA_DT'].dt.date.unique())
                            if not fechas: return 0, 0
                            
                            ciclos = 1
                            inicios_ciclo = [fechas[0]] # Guarda la fecha en que arrancó el ciclo
                            
                            for i in range(1, len(fechas)):
                                # Si la diferencia con el vuelo anterior es mayor a 5 días, es un ciclo NUEVO.
                                # Si es de 5 días o menos, es el MISMO ciclo continuado (Efecto Dron).
                                if (fechas[i] - fechas[i-1]).days > 5:
                                    ciclos += 1
                                    inicios_ciclo.append(fechas[i])
                                    
                            # El intervalo se calcula midiendo los arranques de cada ciclo
                            if ciclos > 1:
                                diffs = [(inicios_ciclo[j] - inicios_ciclo[j-1]).days for j in range(1, ciclos)]
                                avg_int = sum(diffs) / len(diffs)
                            else:
                                avg_int = 0
                                
                            return ciclos, avg_int
                            
                        ciclos_a, int_a = calcular_frecuencia(df_periodo_a)
                        ciclos_b, int_b = calcular_frecuencia(df_periodo_b)
                        
                        c1, c2, c3, c4 = st.columns(4)
                        
                        # 1. Tarjetas de Ciclos (Reemplazamos "Vuelos" por "Ciclos Reales")
                        c1.metric(f"Ciclos Completados ({año_base})", f"{ciclos_a} ciclos")
                        c2.metric(f"Ciclos Completados ({año_comp})", f"{ciclos_b} ciclos", delta=f"{ciclos_b - ciclos_a} ciclos", delta_color="inverse")
                        
                        # 2. Tarjetas de Intervalo
                        str_int_a = f"{int_a:.1f} días" if int_a > 0 else "N/A"
                        str_int_b = f"{int_b:.1f} días" if int_b > 0 else "N/A"
                        c3.metric(f"Intervalo Promedio ({año_base})", str_int_a)
                        
                        if int_a > 0 and int_b > 0:
                            delta_int = int_b - int_a
                            c4.metric(f"Intervalo Promedio ({año_comp})", str_int_b, delta=f"{delta_int:+.1f} días", delta_color="normal")
                        else:
                            c4.metric(f"Intervalo Promedio ({año_comp})", str_int_b)
                        
                        # =====================================================================
                        # =====================================================================
                        # --- 📊 FASE 4: VISORES GRÁFICOS Y ATRIBUCIÓN DE COSTOS ---
                        # =====================================================================
                        st.markdown("---")
                        st.markdown("### 🧬 Análisis de Causa Raíz: Atribución de Variaciones")
                        
                        # ... (código de las variables de avión e insumos) ...
                        # 0. GRÁFICO TENDENCIA TEMPORAL (Evolución Dinámica)
                        st.markdown("#### 📈 Evolución Comparativa: Tendencia del Periodo")
                        
                        # 🎯 CORRECCIÓN: Unimos los datos que YA pasaron por la lupa temporal
                        df_tendencia = pd.concat([df_periodo_a, df_periodo_b])
                        
                        if not df_tendencia.empty:
                            if tipo_periodo in ["AÑO COMPLETO", "POR TRIMESTRE"]:
                                # MODO 1: Zoom de Meses (Para Año o Trimestre)
                                tendencia_agrupa = df_tendencia.groupby(['AÑO', 'MES'])['COSTO_NUM'].mean().reset_index()
                                tendencia_agrupa['EJE_X'] = tendencia_agrupa['MES'].map(meses_dict)
                                tendencia_agrupa = tendencia_agrupa.sort_values('MES')
                                titulo_x = "Meses Operativos"
                            else:
                                # MODO 2: Zoom de Días (Para cuando selecciona un solo mes)
                                df_tendencia['DIA'] = df_tendencia['FECHA_DT'].dt.day
                                tendencia_agrupa = df_tendencia.groupby(['AÑO', 'DIA'])['COSTO_NUM'].mean().reset_index()
                                # Creamos una etiqueta bonita para el día
                                tendencia_agrupa['EJE_X'] = "Día " + tendencia_agrupa['DIA'].astype(str)
                                tendencia_agrupa = tendencia_agrupa.sort_values('DIA')
                                titulo_x = f"Días Operativos ({etiq_periodo})"
                                
                            # Forzamos año a texto para que los colores se separen
                            tendencia_agrupa['AÑO'] = tendencia_agrupa['AÑO'].astype(str)
                            
                            fig_tendencia = px.line(
                                tendencia_agrupa, x='EJE_X', y='COSTO_NUM', color='AÑO', 
                                markers=True, color_discrete_sequence=['#2F75B5', '#ef4444']
                            )
                            
                            # 🎯 CORRECCIÓN DE ETIQUETAS: Claridad absoluta
                            fig_tendencia.update_layout(
                                yaxis_title="Costo Promedio ($ COP / Ha)", 
                                xaxis_title=titulo_x, 
                                plot_bgcolor='rgba(0,0,0,0)',
                                hovermode="x unified"
                            )
                            
                            # Le damos un techo más alto al gráfico para que los números no se corten arriba
                            max_y = tendencia_agrupa['COSTO_NUM'].max() * 1.2
                            fig_tendencia.update_yaxes(range=[0, max_y])

                            fig_tendencia.update_traces(
                                line=dict(width=3), marker=dict(size=8),
                                texttemplate="$ %{y:,.0f}", textposition="top center",
                                hovertemplate="<b>%{x}</b><br>Costo: $ %{y:,.0f} COP/Ha<extra></extra>"
                            )
                            
                            st.plotly_chart(fig_tendencia, use_container_width=True)
                        else:
                            st.warning("⚠️ No hay suficientes operaciones en este periodo exacto para trazar una curva comparativa.")
                            
                        st.markdown("<hr>", unsafe_allow_html=True)
                        
                        # 1. ESCÁNER DE AERONAVE (Específico por Hectárea)
                        col_avion_ha = None
                        for col in df_finca.columns:
                            col_u = str(col).upper().replace('Ó', 'O')
                            # Buscamos estrictamente la tarifa unitaria (que tenga /ha o ha)
                            if 'AVION' in col_u and ('/HA' in col_u or ' HA' in col_u or '(HA)' in col_u):
                                col_avion_ha = col
                                break
                        
                        if col_avion_ha:
                            df_periodo_a['AVION_NUM'] = df_periodo_a[col_avion_ha].apply(convertir_pesos)
                            df_periodo_b['AVION_NUM'] = df_periodo_b[col_avion_ha].apply(convertir_pesos)
                        else:
                            df_periodo_a['AVION_NUM'] = 0.0
                            df_periodo_b['AVION_NUM'] = 0.0

                        # Promedios unitarios reales
                        vuelo_a = df_periodo_a['AVION_NUM'].mean() if not df_periodo_a.empty else 0
                        vuelo_b = df_periodo_b['AVION_NUM'].mean() if not df_periodo_b.empty else 0
                        
                        insumos_a = max(0, costo_a - vuelo_a)
                        insumos_b = max(0, costo_b - vuelo_b)

                        # Escáner Inteligente de Área (Hectáreas) recuperadas
                        col_area = 'AREA_MAESTRA' if 'AREA_MAESTRA' in df_finca.columns else None
                        
                        def limpiar_area(val):
                            try:
                                v = str(val).upper().replace(',', '.')
                                v = "".join([c for c in v if c.isdigit() or c == '.'])
                                return float(v) if v != '' else 0.0
                            except: return 0.0
                            
                        if col_area:
                            df_periodo_a['AREA_NUM'] = df_periodo_a[col_area].apply(limpiar_area)
                            df_periodo_b['AREA_NUM'] = df_periodo_b[col_area].apply(limpiar_area)
                            
                            # Filtro Anti-Clones (Directriz del Comandante)
                            area_a = df_periodo_a.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM'])['AREA_NUM'].sum() if not df_periodo_a.empty else 0
                            area_b = df_periodo_b.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM'])['AREA_NUM'].sum() if not df_periodo_b.empty else 0
                        else:
                            area_a, area_b = 0.0, 0.0

                        # 2. CÁLCULOS GLOBALES INTELIGENTES (Matemática Pura: Tarifa x Hectárea = Inmune a Clones)
                        vuelo_tot_a = vuelo_a * area_a
                        vuelo_tot_b = vuelo_b * area_b
                        insumos_tot_a = insumos_a * area_a
                        insumos_tot_b = insumos_b * area_b

                        # 3. GRÁFICO 1: MATRIZ DE RESPONSABILIDAD (Pestañas Unitario vs Global)
                        st.markdown("#### 🛩️ vs 🧪 Distribución del Encarecimiento")
                        
                        categorias = [f'Análisis {año_base}', f'Análisis {año_comp}']
                        
                        tab_unit, tab_glob = st.tabs(["🎯 Impacto Unitario (Promedio / Ha)", "💰 Impacto Global (Presupuesto Total)"])
                        
                        with tab_unit:
                            fig_unit = go.Figure(data=[
                                go.Bar(name='Costo Avión / Ha', x=categorias, y=[vuelo_a, vuelo_b], marker_color='#2F75B5', text=[f"$ {vuelo_a:,.0f}", f"$ {vuelo_b:,.0f}"], textposition='auto'),
                                go.Bar(name='Costo Insumos (Cóctel) / Ha', x=categorias, y=[insumos_a, insumos_b], marker_color='#548235', text=[f"$ {insumos_a:,.0f}", f"$ {insumos_b:,.0f}"], textposition='auto')
                            ])
                            fig_unit.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor COP / Ha", margin=dict(t=20, b=20))
                            st.plotly_chart(fig_unit, use_container_width=True)
                            
                        with tab_glob:
                            st.markdown("##### 🗺️ Contexto de Área Operada (Efecto Volumen)")
                            g1, g2, g3 = st.columns(3)
                            g1.metric(f"Hectáreas Aplicadas ({año_base})", f"{area_a:,.1f} Ha")
                            g2.metric(f"Hectáreas Aplicadas ({año_comp})", f"{area_b:,.1f} Ha")
                            
                            if area_a > 0:
                                var_area = ((area_b - area_a) / area_a) * 100
                                g3.metric("Variación de Área", f"{var_area:+.1f}%", delta=f"{var_area:+.1f}%", delta_color="off")
                            else:
                                g3.metric("Variación de Área", "N/A")

                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            fig_glob = go.Figure(data=[
                                go.Bar(name='Total Facturación Avión', x=categorias, y=[vuelo_tot_a, vuelo_tot_b], marker_color='#2F75B5', text=[f"$ {vuelo_tot_a:,.0f}", f"$ {vuelo_tot_b:,.0f}"], textposition='auto'),
                                go.Bar(name='Total Consumo Insumos', x=categorias, y=[insumos_tot_a, insumos_tot_b], marker_color='#548235', text=[f"$ {insumos_tot_a:,.0f}", f"$ {insumos_tot_b:,.0f}"], textposition='auto')
                            ])
                            fig_glob.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor Total COP", margin=dict(t=20, b=20))
                            st.plotly_chart(fig_glob, use_container_width=True)
                        
                        # 3. EXPLICACIÓN IA EN TIEMPO REAL (Lógica Corregida y Corporativa)
                        diff_vuelo = vuelo_b - vuelo_a
                        diff_insumos = insumos_b - insumos_a
                        
                        st.info("🧠 **DIAGNÓSTICO AUTOMATIZADO DE IMPACTO:**")
                        if diff_vuelo > 0 and diff_insumos > 0:
                            st.write(f"• La desviación es **MIXTA**: La tarifa operativa (Vuelo) subió **$ {diff_vuelo:,.0f}/Ha** y los insumos (Cóctel) subieron **$ {diff_insumos:,.0f}/Ha**.")
                        elif diff_insumos > 0 and diff_insumos > diff_vuelo:
                            st.write(f"• Factor de mayor impacto: **LOS INSUMOS**. El costo de los químicos generó un alza de **$ {diff_insumos:,.0f}/Ha**, representando el mayor peso en la desviación.")
                        elif diff_vuelo > 0 and diff_vuelo > diff_insumos:
                            st.write(f"• Factor de mayor impacto: **LOGÍSTICA DE VUELO**. La tarifa de aplicación generó un alza de **$ {diff_vuelo:,.0f}/Ha**. Se sugiere revisar tarifas operativas.")
                        elif diff_vuelo <= 0 and diff_insumos <= 0:
                            st.write("• **AHORRO OPERATIVO CONFIRMADO:** Ambos componentes (Vuelo e Insumos) redujeron su costo o se mantuvieron estables en $0 desviación. Excelente control.")
                        else:
                            st.write("• Variación compensada: Las fluctuaciones de vuelo e insumos se equilibraron entre sí.")
                            
                        if area_a > 0 and area_b > 0 and var_area > 5:
                            st.write(f"• **NOTA DE CONTEXTO DE VOLUMEN:** Considere que el área total operada aumentó un **{var_area:.1f}%**, lo cual justifica de forma directa el incremento en el presupuesto global reflejado en la pestaña de impacto total.")
                        # 4. TABLA INTERACTIVA DE CÓCTELES (Apertura Total "Outer Join")
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("#### 📋 Desglose Operativo: Cócteles, Recetas y Volumen Aplicado")
                        
                        col_coctel = 'COCTEL' if 'COCTEL' in df_finca.columns else ('COCTEL_MAESTRO' if 'COCTEL_MAESTRO' in df_finca.columns else None)
                        col_gln = 'GLN_HA' if 'GLN_HA' in df_finca.columns else None
                        
                        if col_coctel:
                            df_periodo_a[col_coctel] = df_periodo_a[col_coctel].astype(str).str.strip().str.upper()
                            df_periodo_b[col_coctel] = df_periodo_b[col_coctel].astype(str).str.strip().str.upper()
                            
                            agg_dict = {'COSTO_NUM': 'mean'}
                            if col_gln: agg_dict[col_gln] = 'mean'
                            
                            g_a = df_periodo_a.groupby(col_coctel).agg(agg_dict).reset_index()
                            g_b = df_periodo_b.groupby(col_coctel).agg(agg_dict).reset_index()
                            
                            # 🎯 LA SOLUCIÓN: Usar "outer" para que aparezcan todos los cócteles y rellenar vacíos con cero
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
                                    f'{col_gln}_BASE': f'Volumen Volado ({año_base}) [Gln/Ha]',
                                    f'{col_gln}_ACTUAL': f'Volumen Volado ({año_comp}) [Gln/Ha]'
                                }, inplace=True)
                                
                                # Etiquetado Inteligente de Novedades
                                def evaluar_dosis(r):
                                    v_base, v_act = float(r[2]), float(r[3])
                                    if v_base == 0: return "⚠️ CÓCTEL NUEVO (No usado año base)"
                                    if v_act == 0: return "⚠️ DESCONTINUADO (No usado año actual)"
                                    if abs(v_act - v_base) > 0.05: return "🚨 CAMBIÓ DOSIS / VOLUMEN"
                                    return "⚖️ MISMA RECETA (Alza SAP)"
                                    
                                tabla_autopsia['Dictamen Dosis'] = tabla_autopsia.apply(evaluar_dosis, axis=1)
                            
                            # Formatear dinero
                            df_vista = tabla_autopsia.copy()
                            df_vista[f'Costo/Ha ({año_base})'] = df_vista[f'Costo/Ha ({año_base})'].map("$ {:,.0f}".format)
                            df_vista[f'Costo/Ha ({año_comp})'] = df_vista[f'Costo/Ha ({año_comp})'].map("$ {:,.0f}".format)
                            df_vista['Variación ($)'] = df_vista['Variación ($)'].map("$ {:,.0f}".format)
                            
                            st.dataframe(df_vista, use_container_width=True)
                        else:
                            st.warning("⚠️ No se encontró la columna 'COCTEL' en la base fusionada para hacer el desglose.")

                    else: st.error("❌ **ERROR DE RADAR:** No se detectó la columna 'FECHA' unificada.")
                else: st.error("❌ **ERROR DE ALINEACIÓN:** No se logró estandarizar Fincas y Costos. Revise encabezados.")
            else: st.error("❌ **ERROR DE VOLUMEN:** Uno de los archivos está vacío.")
# =====================================================================
# --- 🔬 NIVEL 2: AUDITORÍA MOLECULAR (DESGLOSE POR PRODUCTO) ---
# =====================================================================
            st.markdown("<hr>", unsafe_allow_html=True)
            st.error("🔍 REPORTE DE SÓNAR - COLUMNAS VIVAS EN LA MEMORIA:")
                        st.write(df_finca.columns.tolist())
            st.markdown("### 🔬 Nivel 2: Auditoría Molecular de Cócteles (Desglose por Insumo)")
                        
            # 1. Buscamos la columna de Producto/Material (Radar Amplio)
            col_producto = None
            for col in df_finca.columns:
                # Limpiamos cualquier salto de línea basura que envíe Excel
                col_u = str(col).upper().replace('\n', ' ').strip()
                            # Buscamos coincidencias parciales clave
                if 'MATERIAL' in col_u or 'PRODUCTO' in col_u or 'DESCRIPCION' in col_u or 'DESCRIPCIÓN' in col_u or 'INSUMO' in col_u:
                    col_producto = col
                    break
                                
            if col_coctel and col_producto:
                # Recopilamos todos los cócteles que se volaron en el periodo
                cocteles_disponibles = sorted(list(set(df_periodo_a[col_coctel].dropna().unique()) | set(df_periodo_b[col_coctel].dropna().unique())))
                            
                # Selector táctico
                coctel_sel = st.selectbox("🎯 Seleccione un Cóctel para escanear sus componentes químicos:", ["SELECCIONE UN CÓCTEL..."] + cocteles_disponibles)
                            
                if coctel_sel != "SELECCIONE UN CÓCTEL...":
                    # Aislamos las filas de SAP que corresponden SOLO a este cóctel
                    df_coctel_a = df_periodo_a[df_periodo_a[col_coctel] == coctel_sel]
                    df_coctel_b = df_periodo_b[df_periodo_b[col_coctel] == coctel_sel]
                                
                    # Agrupamos por producto para sacar el costo promedio por hectárea de CADA producto
                    prod_a = df_coctel_a.groupby(col_producto)['COSTO_NUM'].mean().reset_index()
                    prod_b = df_coctel_b.groupby(col_producto)['COSTO_NUM'].mean().reset_index()
                                
                    # Unimos la receta de ambos años
                    tabla_molecular = pd.merge(prod_a, prod_b, on=col_producto, how='outer', suffixes=(f'_{año_base}', f'_{año_comp}'))
                    tabla_molecular.fillna(0, inplace=True)
                                
                    # Calculamos al verdadero culpable a nivel molecular
                    tabla_molecular['Variación Costo ($)'] = tabla_molecular[f'COSTO_NUM_{año_comp}'] - tabla_molecular[f'COSTO_NUM_{año_base}']
                                
                    # Formato Ejecutivo
                    tabla_molecular.rename(columns={
                        col_producto: 'INSUMO QUÍMICO / PRODUCTO',
                        f'COSTO_NUM_{año_base}': f'Costo/Ha ({año_base})',
                        f'COSTO_NUM_{año_comp}': f'Costo/Ha ({año_comp})'
                    }, inplace=True)
                                
                    # Ordenamos de mayor a menor variación para que el "culpable" salga de primero
                    tabla_molecular = tabla_molecular.sort_values('Variación Costo ($)', ascending=False)
                                
                    # Aplicamos formato de dinero para visualización
                    df_vista_mol = tabla_molecular.copy()
                    df_vista_mol[f'Costo/Ha ({año_base})'] = df_vista_mol[f'Costo/Ha ({año_base})'].map("$ {:,.0f}".format)
                    df_vista_mol[f'Costo/Ha ({año_comp})'] = df_vista_mol[f'Costo/Ha ({año_comp})'].map("$ {:,.0f}".format)
                    df_vista_mol['Variación Costo ($)'] = df_vista_mol['Variación Costo ($)'].map("$ {:,.0f}".format)
                                
                    st.dataframe(df_vista_mol, use_container_width=True)
            else:
                st.info("💡 Para habilitar el Escáner Molecular, la sábana debe tener una columna llamada 'PRODUCTO', 'MATERIAL' o 'DESCRIPCION'.")
        
        except Exception as e:
            st.error(f"🛰️ **FALLO EN LOS MOTORES:** Error crítico. Motivo: {str(e)}")
