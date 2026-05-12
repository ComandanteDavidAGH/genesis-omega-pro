import pandas as pd
import streamlit as st
import io
import json
import re
import unicodedata
from datetime import datetime
import dateutil.parser

# Imports de conexiones y apis
import openpyxl
import gspread
import plotly.express as px

# Intentar importar matplotlib para el mapa de calor, si falla, el sistema sigue
try:
    import matplotlib
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 2. ARTILLERÍA VISUAL Y CSS ---
arsenal_css = """
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }
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
    st.markdown("<h2 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h2>", unsafe_allow_html=True)
    menu = st.radio("🛰️ SELECCIONE LA OPERACIÓN:", [
        "🏠 Centro de Mando", 
        "🛠️ 1. Mantenimiento Plantilla SAP",
        "📥 2. Carga Facturación", 
        "⚙️ 3. Validación de Misión", 
        "⌨️ 4. Ingreso Manual Acelerado (OS)", 
        "📈 5. Sincronización Precios",
        "✈️ 6. Rastreo Dominicales",
        "⚖️ 7. Arqueo de Inventarios",
        "📊 8. Reporte Hectáreas (Pistas)"
    ])
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
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        f_sabana = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv"], key="sab")
    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        f_pedidos = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv"], key="ped")
    with c3:
        st.markdown("### 🚁 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pis")

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
                                    for r in range(f_h + 1, lim):
                                        fv = str(df.iloc[r, c_idx]).strip()
                                        if fv.lower() in ['nan', '', 'none'] or "TOTAL" in fv.upper(): break
                                        lista_pistas.append({"ORIGEN": f"{f.name} | {n}", "COCTEL": coctel, "FINCA_INFORME": fv, "DATOS_FILA": df.iloc[r].to_dict()})
                                        
                    st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                    st.balloons()
                except Exception as e: st.error(f"🚨 Error: {e}")

# =====================================================================
# ⚙️ 3. VALIDACIÓN DE MISIÓN (NÚCLEO FACTURACIÓN)
# =====================================================================
elif menu == "⚙️ 3. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>Núcleo de Validación y Facturación</h1>", unsafe_allow_html=True)
    
    if 'df_pistas' not in st.session_state or 'df_apoyo' not in st.session_state:
        st.warning("🚨 Cargue los archivos en el Módulo 2 e inicie el procesamiento.")
    else:
        with st.container(border=True):
            st.markdown("### 📡 Panel de Operaciones")
            c0, c1, c2 = st.columns([1, 2, 2])
            fecha_operacion = c0.date_input("📅 Fecha de Vuelo", format="DD/MM/YYYY", key="fecha_vuelo_master")
            
            df_t2 = st.session_state.get('df_config', pd.DataFrame())
            lista_fincas = sorted(df_t2.iloc[:, 0].dropna().unique().tolist()) if not df_t2.empty else []
            finca_sel = c1.selectbox("📍 Seleccione Finca:", ["---"] + lista_fincas)
            
            vuelos_informe = st.session_state['df_pistas']
            vuelo_ref = c2.selectbox("📄 Referencia Pedido/Informe:", ["---"] + vuelos_informe['ORIGEN'].unique().tolist())

        if finca_sel == "---" or vuelo_ref == "---":
            st.info("⚠️ Seleccione Finca y Pedido para rugir motores.")
            st.stop()

        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        finca_limpia = re.sub(r'\s+', ' ', str(finca_sel)).strip().upper()

        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        if not df_t2.empty:
            match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia]
            if not match_t2.empty:
                fila_t2 = match_t2.iloc[0]
                tipo_productor = str(fila_t2.iloc[5]).strip().upper()
                tipo_de_tope_finca = str(fila_t2.iloc[6]).strip().upper()

        # --- 🛰️ EXTRACCIÓN DE INTELIGENCIA DE COSTOS ---
        mult_material = 1.112; tarifa_serv_tec_base = 1337.0; mult_avion_base = 1.112
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
        
        # 🎯 RADAR DE AMPLIO ESPECTRO PARA NÚMEROS DE PEDIDO SAP
        num_pedido = "S/N"
        # Escaneamos desde la columna 18 (T) hasta la 35 (AJ)
        for idx in range(18, 35):
            val_celda = str(datos_raw.get(idx, "")).split('.')[0].strip()
            # Si encontramos un número puro y tiene 7 dígitos o más, es 100% seguro que es el Pedido SAP
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
            ha_dosis_final = r1c3.number_input("🧪 Ha Dosis (Total 459)", value=float(ha_dosis_detectada), key=f"had_{casilla_key}")
            
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
        dict_aviones = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA": 3036525, "AIR TRACTOR": 4665107, "CESSNA ASA": 3666600}
        dict_drones = {"DRONE DATAROT": 84427, "DRONE GENESYS": 75518, "DRONE AVIL": 71280}

        with st.container(border=True):
            st.markdown("#### ✈️ Hangar de Despliegue")
            costo_total_vuelos = 0.0
            total_ha_cobro_escuadron = 0.0

            if mision_solo_dron:
                st.success("🚁 Modo Dron Activo: Costos calculados sin recargos terrestres ni topes de pista.")
                df_drones_def = pd.DataFrame([{"Drone": "DRONE DATAROT", "Hectáreas": float(ha_cobro_detectada)}])
                escuadron_drones = st.data_editor(
                    df_drones_def, key=f"drones_{casilla_key}", num_rows="dynamic",
                    column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True
                )
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    costo_total_vuelos += (dict_drones.get(dr_sel, 0) * ha_dr) * mult_avion_final

            else:
                c_av, c_dr = st.columns(2)
                with c_av:
                    st.markdown("##### 🛩️ Base Aviones")
                    df_aviones_def = pd.DataFrame([{"Avión": "THRUS SR2", "Hectáreas": float(ha_cobro_detectada), "Horómetro": 1.00}])
                    escuadron_aviones = st.data_editor(df_aviones_def, key=f"aviones_{casilla_key}", num_rows="dynamic", column_config={"Avión": st.column_config.SelectboxColumn("Modelo", options=list(dict_aviones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f"), "Horómetro": st.column_config.NumberColumn("Horómetro", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)
                with c_dr:
                    st.markdown("##### 🚁 Base Drones (Apoyo)")
                    df_drones_def = pd.DataFrame([{"Drone": None, "Hectáreas": 0.0}])
                    escuadron_drones = st.data_editor(df_drones_def, key=f"drones_mix_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys())), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)
                
                for _, row in escuadron_aviones.iterrows():
                    av_sel, ha_av, horo = row["Avión"], float(row.get("Hectáreas", 0)), float(row.get("Horómetro", 0))
                    if pd.isna(av_sel) or ha_av <= 0: continue
                    total_ha_cobro_escuadron += ha_av
                    tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av
                    tarifa_aplicada = tarifa_base_ha + recargo_final if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope) + recargo_final
                    costo_total_vuelos += (tarifa_aplicada * ha_av) * mult_avion_final
                
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    costo_total_vuelos += (dict_drones.get(dr_sel, 0) * ha_dr) * mult_avion_final

        st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Mezcla")
        costo_mezcla_total = 0.0

        if not match_ped.empty:
            idx_precio = -1; idx_lote = -1; idx_saldo = -1
            if not df_sab.empty:
                for j, col in enumerate(df_sab.columns):
                    col_str = str(col).upper()
                    if 'MAYOR' in col_str or 'PRECIO' in col_str: idx_precio = j
                    if 'LOTE' in col_str: idx_lote = j
                    if ('LIBRE' in col_str or 'SALDO' in col_str) and 'VALOR' not in col_str: idx_saldo = j

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
                sap_dict_pista[nombre_limpio] = dosis_pista
                datos_extraidos_sap.append({"cod": cod_item, "nombre": nombre_p, "nombre_limpio": nombre_limpio, "cant_total": cant_total})

            # ====================================================================
            # 🧠 MOTOR IA MAESTRO: LÓGICA DE LA MACRO ORIGINAL (DD_Mesclas)
            # ====================================================================
            dict_recetas = {}
            dict_lideres = {}
            dict_fertilizantes = {}

            if not df_mez.empty:
                for idx, row in df_mez.iterrows():
                    # 1. Leer Recetas Base (Columnas A, B, C, D)
                    if len(row) > 3:
                        cid = str(row.iloc[0]).strip().upper()
                        p_tabla_clean = str(row.iloc[1]).strip().upper().replace(" ", "")
                        d_tabla = extraer_numero(row.iloc[2])
                        es_lider = str(row.iloc[3]).strip().upper() == "X"

                        if cid and cid != 'NAN' and p_tabla_clean:
                            if cid not in dict_recetas: dict_recetas[cid] = {}
                            dict_recetas[cid][p_tabla_clean] = d_tabla
                            if es_lider: dict_lideres[cid] = p_tabla_clean

                    # 2. Leer Tabla de Fertilizantes (Columnas M y N)
                    if len(row) > 13:
                        fert_name = str(row.iloc[12]).strip().upper()
                        fert_sigla = str(row.iloc[13]).strip().upper()
                        if fert_name and fert_sigla and fert_name not in ["NAN", "FERTILIZANTES", ""]:
                            dict_fertilizantes[fert_name.replace(" ", "")] = fert_sigla

            # --- FASE 1: BUSCAR EL CÓCTEL BASE (CRUCE POR PRODUCTO LÍDER) ---
            coctel_base = "SIN COINCIDENCIA"
            dosis_oficiales_coctel = {}
            max_p = -999

            for iter_id, receta in dict_recetas.items():
                es_valido = True
                puntaje = 0
                lider_db = dict_lideres.get(iter_id, "")

                # Validar que la 'X' (Producto Líder) esté presente en los materiales de SAP
                match_lider = False
                if lider_db:
                    for k_sap in sap_dict_pista.keys():
                        if lider_db == k_sap or (len(k_sap)>=4 and lider_db in k_sap) or (len(lider_db)>=4 and k_sap in lider_db):
                            match_lider = True; break

                if match_lider: puntaje += 1000
                else: es_valido = False

                # Validar los demás productos para dar con la receta exacta
                if es_valido:
                    for p_receta, d_esperada in receta.items():
                        match_receta = False
                        dose_matched = False
                        for k_sap, d_sap in sap_dict_pista.items():
                            if p_receta == k_sap or (len(k_sap)>=4 and p_receta in k_sap) or (len(p_receta)>=4 and k_sap in p_receta):
                                match_receta = True
                                if abs(d_sap - d_esperada) <= 0.2: dose_matched = True
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
                        sigla_fertilizante = f"+{f_sigla}"
                        break
                if sigla_fertilizante: break

            coctel_ganador = coctel_base + sigla_fertilizante if coctel_base != "SIN COINCIDENCIA" else "SIN COINCIDENCIA"

            st.success(f"🤖 **MOTOR IA MAESTRO:** Cóctel Oficial: **{coctel_ganador}** (Base de Datos: {coctel_base} | Fertilizante: {sigla_fertilizante or 'Ninguno'})")

            # --- FASE 3: CONSTRUIR LA MATRIZ DE LIQUIDACIÓN ---
            matriz_datos = []
            for item_data in datos_extraidos_sap:
                cod_item = item_data['cod']
                nombre_p = item_data['nombre']
                nombre_limpio = item_data['nombre_limpio']
                cant_total_pedido = item_data['cant_total']

                costo_unit = 0.0; lote_sap = "SIN LOTE EN PISTA"; saldo_sap = 0.0

                if not df_sab.empty:
                    match_sabana_global = df_sab[df_sab.iloc[:, 0].astype(str).str.strip() == cod_item]
                    if match_sabana_global.empty: match_sabana_global = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(cod_item, case=False, na=False)).any(axis=1)]

                    if not match_sabana_global.empty:
                        fila_precio = match_sabana_global.iloc[0]
                        if idx_precio != -1: costo_unit = extraer_numero(fila_precio.iloc[idx_precio])
                        if costo_unit == 0.0:
                            col_valor_tot = [c for c in fila_precio.index if 'VALOR' in str(c).upper() and 'LIBRE' in str(c).upper()]
                            col_cant_tot = [c for c in fila_precio.index if 'LIBRE' in str(c).upper() and 'VALOR' not in str(c).upper()]
                            if col_valor_tot and col_cant_tot:
                                v_total = extraer_numero(fila_precio[col_valor_tot[0]])
                                c_total = extraer_numero(fila_precio[col_cant_tot[0]])
                                if c_total > 0: costo_unit = v_total / c_total

                        match_pista = match_sabana_global[match_sabana_global.astype(str).apply(lambda x: x.str.contains(pista_sel, case=False, na=False)).any(axis=1)]
                        if not match_pista.empty:
                            try:
                                col_ordenar = [c for c in match_pista.columns if ('LIBRE' in str(c).upper() or 'SALDO' in str(c).upper()) and 'VALOR' not in str(c).upper()]
                                if col_ordenar:
                                    match_pista['Temp_Sort'] = match_pista[col_ordenar[0]].apply(extraer_numero)
                                    match_pista = match_pista.sort_values(by='Temp_Sort', ascending=False)
                            except: pass
                            fila_pista = match_pista.iloc[0]
                            if idx_lote != -1: lote_sap = str(fila_pista.iloc[idx_lote])
                            if idx_saldo != -1: saldo_sap = extraer_numero(fila_pista.iloc[idx_saldo])

                # 🎯 ASIGNAR LA DOSIS (1. De la Receta, 2. Rescate Matemático)
                dosis_teorica = None

                # Buscar en la receta base oficial
                for p_receta, d_oficial in dosis_oficiales_coctel.items():
                    if p_receta == nombre_limpio or (len(nombre_limpio)>=4 and p_receta in nombre_limpio) or (len(p_receta)>=4 and nombre_limpio in p_receta):
                        dosis_teorica = d_oficial
                        break

                # 🚀 REGLA DE ORO EXCEL: SOBREESCRITURA CONDICIONAL MULTIPLE
                if "ACONDICIONADOR" in nombre_limpio:
                    if "ZN" in coctel_ganador or "BT" in coctel_ganador:
                        dosis_teorica = 0.06
                    else:
                        dosis_teorica = 0.02
                elif "IMBIOSIL" in nombre_limpio.replace(" ", "") or "INBIOMAG" in nombre_limpio:
                    # Si el cóctel empieza con "IN", es el líder (1.5). Si va de mezcla, es 1.0.
                    if coctel_ganador.startswith("IN"):
                        dosis_teorica = 1.5
                    else:
                        dosis_teorica = 1.0
                
                # Rescate final (Si es un fertilizante o Mancol, se calcula matemáticamente)
                if dosis_teorica is None:
                    dosis_teorica = cant_total_pedido / ha_dosis_final if ha_dosis_final > 0 else 0.0

                # 🚀 BLINDAJE SAP: Redondear el costo unitario a 0 decimales antes de multiplicar
                costo_margen = round(costo_unit * mult_material, 0)

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
            # 🚀 ESTA ES LA LÍNEA QUE SE BORRÓ O QUEDÓ MAL TABULADA:
            df_matriz = pd.DataFrame(matriz_datos)
                
            if 'editor_valid' in st.session_state:
                ediciones = st.session_state['editor_valid'].get('edited_rows', {})
                for row_idx, edit_dict in ediciones.items():
                    if "B: Dosis/Ha (SAP)" in edit_dict: df_matriz.at[row_idx, "B: Dosis/Ha (SAP)"] = edit_dict["B: Dosis/Ha (SAP)"]
                    if "C: X (Extra %)" in edit_dict: df_matriz.at[row_idx, "C: X (Extra %)"] = edit_dict["C: X (Extra %)"]

            df_matriz["B_Val"] = df_matriz["B: Dosis/Ha (SAP)"].fillna(0.0)
            df_matriz["C_Val"] = df_matriz["C: X (Extra %)"].fillna(0.0)
            df_matriz["D: Dosis Total (Sistema)"] = (df_matriz["B_Val"] * (1 + df_matriz["C_Val"]/100) * ha_dosis_final).round(3)

            costo_mezcla_total = (df_matriz["D: Dosis Total (Sistema)"] * df_matriz["E: Costo Unit (+Margen)"]).sum()
            df_matriz = df_matriz.drop(columns=["B_Val", "C_Val"])

            edited_df = st.data_editor(
                df_matriz, key='editor_valid',
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
        # ====================================================================
        # 💰 LIQUIDACIÓN FINAL Y CAJAS DE COPIA SAP (TRAZABILIDAD FINCA A FINCA)
        # ====================================================================
        st.markdown("---")
        st.markdown("### 💰 Liquidación Final (Bóveda SAP)")
        
        # 1. Cálculos Unitarios Puros (Redondeados a cero decimales como SAP)
        unitario_st = round(d_ciclo_factura * tarifa_serv_tec_base, 0)
        unitario_vuelo = round(costo_total_vuelos / total_ha_cobro_escuadron if total_ha_cobro_escuadron > 0 else 0, 0)
        
        # 2. 🎯 CÁLCULOS TOTALES EXCLUSIVOS DE ESTA FINCA (Para Trazabilidad SAP)
        subtotal_st_finca = round(unitario_st * ha_dosis_final, 0)
        subtotal_vuelo_finca = round(unitario_vuelo * ha_dosis_final, 0)
        
        gran_total = costo_mezcla_total + subtotal_vuelo_finca + subtotal_st_finca
        costo_por_ha = gran_total / ha_dosis_final if ha_dosis_final > 0 else 0

        # --- MÉTRICAS DE CONTROL ---
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("🚜 Hectáreas Factura (Finca)", f"{ha_dosis_final:.2f} Ha")
        
        if mision_solo_dron: 
            r2.metric("🛣️ Condición Pista", "NO APLICA (Dron)")
        else: 
            r2.metric("🛣️ Condición Pista", tipo_de_tope_finca, f"Límite: $ {fmt_sap(val_tope)}")
            
        r3.metric("👨‍🔬 Tarifa ST Base", f"$ {fmt_sap(tarifa_serv_tec_base)}")
        r4.metric("✈️ Multiplicador Avión", f"x {mult_avion_final}")

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- 📋 CAJAS DE COPIA RÁPIDA (UNITARIOS PARA SAP) ---
        st.markdown("#### 📋 Cajas de Copia para Digitación en SAP")
        c_sap1, c_sap2, c_sap3, c_sap4 = st.columns(4)
        
        with c_sap1: 
            st.caption("👨‍🔬 UNITARIO Serv. Tec (Pos. 459)") # <- Corregida la etiqueta para que coincida con SAP
            st.code(fmt_sap(unitario_vuelo), language="text") # El valor de 73.710 va a la 459 en SAP
            
        with c_sap2: 
            st.caption("✈️ UNITARIO Vuelo (Pos. 429)") # <- Corregida la etiqueta para que coincida con SAP
            st.code(fmt_sap(unitario_st), language="text") # El valor de 52.850 va a la 429 en SAP

        with c_sap3: 
            st.caption("🧪 TOTAL Mezcla Química")
            st.code(fmt_sap(costo_mezcla_total), language="text")
            
        with c_sap4:
            st.markdown(f"""
            <div style='background-color:#0d1b2a; padding:10px; border-radius:5px; border:1px solid #d4af37; text-align:center;'>
                <p style='margin:0; color:#d4af37; font-size:12px;'>💰 COSTO x HECTÁREA (Final)</p>
                <h4 style='margin:0; color:white;'>$ {fmt_sap(costo_por_ha)}</h4>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- TOTALES INFORMATIVOS ---
        st.markdown("##### 💵 Totales de Posiciones por Finca (Informativo)")
        c_tot1, c_tot2, c_tot3 = st.columns(3)
        c_tot1.metric("Subtotal Serv. Tec (459)", f"$ {fmt_sap(subtotal_vuelo_finca)}") # Etiquetas invertidas para igualar a su pantalla
        c_tot2.metric("Subtotal Vuelo (429)", f"$ {fmt_sap(subtotal_st_finca)}")
        c_tot3.metric("🔥 GRAN TOTAL FINCA", f"$ {fmt_sap(gran_total)}")
        
        # ====================================================================
        # 🛰️ COORDENADAS DE LANZAMIENTO (NO BORRAR)
        # ====================================================================
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
            with st.spinner("🚀 Inyectando datos en TABLA 1 y APOYO..."):
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

                    ha_f = float(ha_dosis_final)
                    h_total_v = (ha_f / 10) if mision_solo_dron else 1.0
                    vol_total_gln = ha_f * 6
                    rend_min = h_total_v * 60
                    piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                    hk_f = "DR51" if "DATAROT" in tipo_mision else "DR52" if "GENESYS" in tipo_mision else "DR53" if "AVIL" in tipo_mision else "S/N"

                    row_azul = [""] * 34
                    row_azul[0] = os_virtual
                    row_azul[1] = bloque_f
                    row_azul[2] = finca_limpia
                    row_azul[3] = sector_f
                    row_azul[4] = ha_bruta_f
                    row_azul[5] = ha_f
                    row_azul[6] = coctel_ganador
                    row_azul[7] = fecha_str
                    row_azul[8] = dia_sem
                    row_azul[9] = num_sem
                    row_azul[10] = h_total_v
                    row_azul[11] = 6
                    row_azul[12] = round(vol_total_gln, 2)
                    row_azul[13] = round(h_total_v, 2)
                    row_azul[14] = round(rend_min, 2)
                    row_azul[15] = piloto_f
                    row_azul[16] = hk_f
                    row_azul[17] = tipo_mision
                    row_azul[18] = float(gran_total)
                    row_azul[19] = float(costo_por_ha)
                    row_azul[20] = float(recargo_final)
                    row_azul[21] = float(gran_total)
                    row_azul[23] = pista_manual
                    row_azul[28] = float(gran_total)
                    row_azul[32] = tipo_productor
                    row_azul[33] = "GÉNESIS_V2_PRO"

                    fila_apoyo = [""] * 15
                    fila_apoyo[0] = "=IFERROR(ROW()-3; 0)" 
                    fila_apoyo[1] = finca_limpia
                    fila_apoyo[2] = ha_f
                    fila_apoyo[3] = float(costo_por_ha)
                    fila_apoyo[5] = fecha_str
                    fila_apoyo[8] = coctel_ganador
                    fila_apoyo[10] = pista_manual
                    fila_apoyo[13] = tipo_mision
                    
                    hoja_maestra.append_row(row_azul, value_input_option='USER_ENTERED')
                    hoja_apoyo.append_row(fila_apoyo, value_input_option='USER_ENTERED')

                    st.balloons()
                    st.success(f"✅ IMPACTO TOTAL CONFIRMADO. Referencia: {os_virtual}")
                    
                    if 'memoria_excel' in st.session_state:
                        del st.session_state['memoria_excel']

                except Exception as e_save:
                    st.error(f"🚨 Falla en el Gatillo de Guardado: {e_save}")

# =====================================================================
# ⌨️ 4. INGRESO MANUAL ACELERADO (ESTÁ ARRIBA EN EL ORDEN LÓGICO)
# =====================================================================

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
    
    st.title("⚖️ Arqueo de Inventarios y Conciliación")
    
    archivo_sap = st.sidebar.file_uploader("1️⃣ Sábana de SAP", type=['xlsx', 'csv'])
    archivos_sup = st.sidebar.file_uploader("2️⃣ Reportes Supervisores (.xlsx)", type=['xlsx'], accept_multiple_files=True)
    semana_obj = st.sidebar.text_input("🎯 Semana a Auditar (Ej: 17):", placeholder="Escriba aquí...")

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

    if st.sidebar.button("🚀 INICIAR ARQUEO ESTRATÉGICO", use_container_width=True):
        if not archivo_sap or not archivos_sup or not semana_obj:
            st.sidebar.error("❌ Faltan suministros.")
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
                        st.error("No se encontraron datos.")

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
