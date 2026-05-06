import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import gspread
import plotly.express as px

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 2. ARTILLERÍA VISUAL ---
arsenal_css = """
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }
</style>
"""
st.markdown(arsenal_css, unsafe_allow_html=True)

# --- 3. MENÚ TÁCTICO ---
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h2>", unsafe_allow_html=True)
    menu = st.radio("🛰️ NAVEGACIÓN:", ["🏠 Centro de Mando", "📥 1. Buzón de Carga", "⚙️ 2. Validación de Misión", "📊 3. Arqueo y Reportes", "🛡️ Configuración"])
    st.info(f"📅 Operación: {datetime.now().strftime('%Y-%m-%d')}")

# --- 4. LÓGICA DE CARGA ---

if menu == "🏠 Centro de Mando":
    st.markdown("<h1 class='titulo-principal'>Centro de Mando</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3>Estrategia de Validación (La Trinidad):</h3>
        <ol>
            <li><b>Sábana SAP:</b> Validamos Lotes y Precios oficiales.</li>
            <li><b>Pedidos SAP:</b> Validamos lo que se DEBÍA hacer (Fincas/Hectáreas).</li>
            <li><b>Informes Pista:</b> Validamos lo que REALMENTE se hizo.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

elif menu == "📥 1. Buzón de Carga":
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Cuartel General</h1>", unsafe_allow_html=True)
    
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
            with st.spinner("Sincronizando los 3 frentes y conectando con Satélite en Google Drive..."):
                try: 
                    # 1. Leer Sábana
                    bytes_sabana = io.BytesIO(f_sabana.getvalue())
                    if f_sabana.name.lower().endswith(('.xlsx', '.xls')):
                        st.session_state['df_sabana'] = pd.read_excel(bytes_sabana)
                    else:
                        st.session_state['df_sabana'] = pd.read_csv(bytes_sabana, sep=None, engine='python')
                    
                    # 2. Leer Pedidos
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    if f_pedidos.name.lower().endswith(('.xlsx', '.xls')):
                        st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos)
                    else:
                        st.session_state['df_pedidos'] = pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    # 3. CONEXIÓN SATELITAL
                    try:
                        if "gcp_credentials" in st.secrets:
                            cred_dict = dict(st.secrets["gcp_credentials"])
                            gc = gspread.service_account_from_dict(cred_dict)
                        else:
                            gc = gspread.service_account(filename='credenciales.json')
                        
                        url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                        boveda = gc.open_by_url(url_boveda)
                        
                        # Cargar Configuración
                        hoja_tabla2 = boveda.worksheet("TABLA 2")
                        datos_tabla2 = hoja_tabla2.get_all_values()
                        st.session_state['df_config'] = pd.DataFrame(datos_tabla2[1:], columns=datos_tabla2[0])
                        
                        # Cargar Tabla de Apoyo
                        hoja_apoyo = boveda.worksheet("TablaNegraDatos") 
                        datos_apoyo = hoja_apoyo.get_all_values()
                        st.session_state['df_apoyo'] = pd.DataFrame(datos_apoyo[1:], columns=datos_apoyo[0])
                        
                    except Exception as error_nube:
                        st.error(f"🚨 Falla en el Enlace Satelital: {error_nube}")
                    
                    # 4. ESCÁNER DE BARRIDO TOTAL
                    lista_pistas = []
                    for f in f_pistas:
                        try:
                            bytes_data = f.getvalue()
                            ext = f.name.split('.')[-1].lower()
                            
                            if ext == 'xlsx':
                                wb = openpyxl.load_workbook(io.BytesIO(bytes_data), read_only=True)
                                visibles = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
                                wb.close()
                                dict_pestañas = pd.read_excel(io.BytesIO(bytes_data), sheet_name=visibles, header=None)
                            else:
                                dict_pestañas = pd.read_excel(io.BytesIO(bytes_data), sheet_name=None, header=None)

                            for nombre_pestaña, df in dict_pestañas.items():
                                df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                                filas_coctel = df[df.astype(str).apply(lambda x: x.str.contains('COCTEL', case=False, na=False)).any(axis=1)].index.tolist()
                                
                                for i_c, fila_c_idx in enumerate(filas_coctel):
                                    fila_data = df.iloc[fila_c_idx].dropna().tolist()
                                    nombre_coctel = fila_data[1] if len(fila_data) > 1 else "DESCONOCIDO"
                                    limite_inferior = filas_coctel[i_c + 1] if i_c + 1 < len(filas_coctel) else len(df)
                                    df_segmento = df.iloc[fila_c_idx:limite_inferior]
                                    idx_fincas = df_segmento[df_segmento.astype(str).apply(lambda x: x.str.contains('FINCAS', case=False, na=False)).any(axis=1)].index
                                    
                                    if not idx_fincas.empty:
                                        fila_h_fincas = idx_fincas[0]
                                        col_fincas_idx = (df.iloc[fila_h_fincas].astype(str).str.contains('FINCAS', case=False)).values.argmax()
                                        for r in range(fila_h_fincas + 1, limite_inferior):
                                            finca_v = str(df.iloc[r, col_fincas_idx]).strip()
                                            if finca_v.lower() in ['nan', '', 'none'] or "TOTAL" in finca_v.upper():
                                                break
                                            lista_pistas.append({
                                                "ORIGEN": f"{f.name} | {nombre_pestaña}",
                                                "COCTEL": nombre_coctel,
                                                "FINCA_INFORME": finca_v,
                                                "DATOS_FILA": df.iloc[r].to_dict()
                                            })
                        except Exception as e_file:
                            st.error(f"🚨 Error en archivo {f.name}: {e_file}")

                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                        st.success(f"✅ ¡Barrido Exitoso! {len(lista_pistas)} vuelos detectados.")
                    else:
                        st.error("🚨 No se encontró la estructura de 'FINCAS'.")

                except Exception as e_master: 
                    st.error(f"🚨 Error Crítico en Procesamiento: {e_master}")

elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>🚀 Módulo 2: Núcleo de Validación y Facturación</h1>", unsafe_allow_html=True)
    
    # --- RADAR DE DIAGNÓSTICO (NUEVO BLINDAJE) ---
    faltan_datos = False
    if 'df_pistas' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró el Informe de Pistas en la memoria. ¿Presionó el botón de procesar en el Módulo 1?")
        faltan_datos = True
    if 'df_apoyo' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró la 'TablaNegraDatos' de Google Drive. Revise la conexión satelital en el Módulo 1.")
        faltan_datos = True
        
    if faltan_datos:
        st.warning("⚠️ Vuelva al Módulo 1, cargue los archivos, presione INICIAR PROCESAMIENTO y espere el mensaje verde de éxito antes de volver aquí.")
    else:
        # --- 1. UBICACIÓN ESTRATÉGICA DE DATOS GENERALES ---
        with st.container(border=True):
            st.markdown("### 📡 Panel de Operaciones (Datos de Vuelo)")
            c1, c2 = st.columns(2)
            
            # LISTA DEPLEGABLE DE FINCAS (Desde TablaNegraDatos / TABLA DE APOYO 2023)
            lista_fincas_apoyo = st.session_state['df_apoyo'].iloc[:, 0].unique().tolist() # Asumiendo columna 0 es Finca
            finca_sel = c1.selectbox("📍 Seleccione Finca (Base de Datos):", ["---"] + lista_fincas_apoyo)
            
            # Lista de vuelos detectados en informes para cruzar pedido
            vuelos_informe = st.session_state['df_pistas']
            vuelo_ref = c2.selectbox("📄 Referencia Pedido/Informe:", ["---"] + vuelos_informe['ORIGEN'].unique().tolist())

        if finca_sel != "---" and vuelo_ref != "---":
            # Extraer datos de los informes cargados
            datos_vuelo = vuelos_informe[vuelos_informe['ORIGEN'] == vuelo_ref].iloc[0]
            datos_raw = datos_vuelo['DATOS_FILA'] # Diccionario de la fila de Excel
            
            # --- 2. CONFIGURACIÓN DE AERONAVE Y PISTA ---
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns(4)
                
                # Aeronave (Manual)
                lista_aviones = ["THRUS SR2", "PIPER PA 36-375", "CESSNA O PIPER PA", "AIR TRACTOR", "CESSNA ASA", "DRONE DATAROT", "DRONE GENESYS", "DRONE AVIL"]
                avion_sel = c1.selectbox("✈️ Tipo de Avión:", lista_aviones)
                
                # Pista (Dual: Lista + Valor automático de Pedido)
                pista_pedido = str(datos_raw.get(2, "PLUC")) # Intenta traer pista del pedido
                lista_pistas = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
                idx_pista = lista_pistas.index(pista_pedido) if pista_pedido in lista_pistas else 0
                pista_sel = c2.selectbox("🛣️ Pista Operativa:", lista_pistas, index=idx_pista)
                
                # Horómetro y Hectáreas
                ha_inf = float(datos_raw.get(8, 0)) # Columna Hectáreas del informe
                horo = c3.number_input("⏱️ Horómetro:", value=1.00, step=0.01)
                ha_real = c4.number_input("🚜 Hectáreas Reales:", value=ha_inf, step=0.1)

            # --- 3. CÁLCULOS AUTOMÁTICOS DE TARIFA (ESPEJO EXCEL) ---
            dict_precios = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA": 3036525, "AIR TRACTOR": 4665107, "CESSNA ASA": 3666600, "DRONE DATAROT": 84427, "DRONE GENESYS": 75518, "DRONE AVIL": 71280}
            dict_topes = {"PLUC": "TOPE MAX GENERAL", "PORI": "TOPE SUR", "PDIV": "TOPE PARCELA INTER < 20ha", "TEHO": "TOPE MAX GENERAL", "LUCI": "TOPE SUR"}
            
            valor_base = dict_precios.get(avion_sel, 0)
            tope_msj = dict_topes.get(pista_sel, "SIN TOPE")
            
            # Lógica Recargo DIVAS/PDIV (Dual)
            recargo_terrestre = 0
            if (pista_sel == "PDIV" or pista_sel == "LUCI") and ("DRONE" not in avion_sel):
                recargo_terrestre = 45000 # Valor porción terrestre para aviones
            
            with st.container(border=True):
                st.markdown("#### 💰 Liquidación de Vuelo")
                m1, m2, m3 = st.columns(3)
                m1.metric("Precio Base (Hora/Ha)", f"${valor_base:,.0f}")
                m2.metric("Tope de Pista", tope_msj)
                m3.metric("Recargo Terrestre (DIVAS)", f"${recargo_terrestre:,.0f}")

            # --- 4. GRAN MATRIZ DE PRODUCTOS (COLUMNAS A A I) ---
            st.markdown("#### 🧪 Matriz de Validación de Productos y Dosis")
            
            # Aquí el sistema cruza Pedidos SAP, Informe Pista y Sábana SAP
            matriz_datos = []
            
            # Escaneamos los productos que reportó el supervisor en el informe
            for i in range(10, 18):
                consumo_sup = datos_raw.get(i, 0)
                if pd.notnull(consumo_sup) and float(consumo_sup) > 0:
                    nombre_prod = f"Producto_{i}" # Aquí buscaremos el nombre real en Pedidos SAP
                    dosis_base = 0.5 # Aquí traemos dosis de SAP
                    porcentaje_x = 1.0 # Columna C (La X)
                    
                    matriz_datos.append({
                        "A: Producto": nombre_prod,
                        "B: Dosis SAP": dosis_base,
                        "C: X (%)": porcentaje_x,
                        "D: Dosis/Ha": dosis_base * porcentaje_x,
                        "E: Costo (Margen)": 15000, # Cruzado con Tipo Productor
                        "G: Lotes": "L-2024-X", # Desde Sábana SAP
                        "H: Saldo Real SAP": 500, # Desde Sábana SAP
                        "I: Consumo Supervisor": float(consumo_sup)
                    })
            
            if matriz_datos:
                df_final = pd.DataFrame(matriz_datos)
                st.dataframe(df_final, use_container_width=True)
            
            # --- BOTÓN DE CIERRE ---
            if st.button("🔥 PROCESAR LIQUIDACIÓN FINAL", type="primary", use_container_width=True):
                st.balloons()
                st.success(f"Liquidación de {finca_sel} procesada exitosamente.")
