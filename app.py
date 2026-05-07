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
                        hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023") 
                        datos_apoyo = hoja_apoyo.get_all_values()
                        st.session_state['df_apoyo'] = pd.DataFrame(datos_apoyo[1:], columns=datos_apoyo[0])
                        hoja_mezclas = boveda.worksheet("DD_Mesclas")
                        st.session_state['df_mezclas'] = pd.DataFrame(hoja_mezclas.get_all_values())
                        
                        hoja_conf = boveda.worksheet("Configuración")
                        st.session_state['df_config_base'] = pd.DataFrame(hoja_conf.get_all_values())
                        # ------------------------------------------------------
                    
                        
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
    
    # --- RADAR DE DIAGNÓSTICO ---
    faltan_datos = False
    if 'df_pistas' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró el Informe de Pistas. Vaya al Módulo 1.")
        faltan_datos = True
    if 'df_apoyo' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró la 'TABLA DE APOYO2023'. Vaya al Módulo 1.")
        faltan_datos = True
        
    if faltan_datos:
        st.warning("⚠️ Vuelva al Módulo 1, cargue los archivos y presione INICIAR PROCESAMIENTO.")
    else:
        # --- 1. UBICACIÓN ESTRATÉGICA ---
        with st.container(border=True):
            st.markdown("### 📡 Panel de Operaciones (Datos de Vuelo)")
            c1, c2 = st.columns(2)
            
            lista_fincas_apoyo = st.session_state['df_apoyo'].iloc[:, 1].unique().tolist() if 'df_apoyo' in st.session_state else []
            finca_sel = c1.selectbox("📍 Seleccione Finca (Base de Datos):", ["---"] + lista_fincas_apoyo)
            
            vuelos_informe = st.session_state.get('df_pistas', pd.DataFrame())
            lista_vuelos = vuelos_informe['ORIGEN'].unique().tolist() if not vuelos_informe.empty else []
            vuelo_ref = c2.selectbox("📄 Referencia Pedido/Informe:", ["---"] + lista_vuelos)

        # 🛑 FRENO DE EMERGENCIA
        if finca_sel == "---" or vuelo_ref == "---":
            st.info("⚠️ Seleccione una Finca y un Informe/Pedido para iniciar la liquidación.")
            st.stop() 
            
        # =====================================================================================
        # 🟢 NÚCLEO DE PROCESAMIENTO
        # =====================================================================================
        import re
        def extraer_numero(valor):
            if pd.isna(valor) or valor == "": return 0.0
            if isinstance(valor, (int, float)): return float(valor)
            v = str(valor).strip().upper()
            v = re.sub(r'[^\d.,-]', '', v) 
            if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
            elif ',' in v: v = v.replace(',', '.') 
            try: return float(v)
            except: return 0.0

        datos_vuelo = vuelos_informe[vuelos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo['DATOS_FILA'] 
        
        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame()) 
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        # --- A. EXTRACCIÓN DE IDENTIDAD DE FINCA ---
        tipo_productor = "SOCIO"
        if not df_apoyo.empty:
            match_finca = df_apoyo[df_apoyo.iloc[:, 1].astype(str).str.contains(str(finca_sel), case=False, na=False)]
            if not match_finca.empty:
                col_tipo = [c for c in df_apoyo.columns if 'TIPO' in str(c).upper() or 'GRUPO' in str(c).upper()]
                if col_tipo: tipo_productor = str(match_finca.iloc[0][col_tipo[0]]).upper()

        # --- B. LECTURA EXACTA DE TABLA DE CONFIGURACIÓN (Multiplicadores) ---
        mult_material = 1.112; tarifa_serv_tec = 1337.0; mult_avion = 1.112
        if not df_cfg.empty:
            match_prod = df_cfg[df_cfg.iloc[:, 0].astype(str).str.contains(tipo_productor, case=False, na=False)]
            if not match_prod.empty:
                fila_cfg = match_prod.iloc[0]
                if len(fila_cfg) >= 6: # Según las columnas de su imagen
                    mult_material = extraer_numero(fila_cfg.iloc[2]) # Columna Material (Ej: 1,112)
                    tarifa_serv_tec = extraer_numero(fila_cfg.iloc[3]) # Columna Servicio Tec (Ej: 1337)
                    mult_avion = extraer_numero(fila_cfg.iloc[5]) # Columna Avion (Ej: 1,112)

        # --- C. MOTOR DÍAS CICLO ---
        dias_ciclo_calc = 0
        if not df_apoyo.empty:
            col_fecha_hist = [c for c in df_apoyo.columns if 'FECHA' in str(c).upper()]
            if col_fecha_hist:
                hist_finca = df_apoyo[df_apoyo.iloc[:, 1].astype(str).str.contains(str(finca_sel), case=False, na=False)].copy()
                if not hist_finca.empty:
                    hist_finca['FECHA_DT'] = pd.to_datetime(hist_finca[col_fecha_hist[0]], errors='coerce')
                    fecha_vuelo_actual = pd.to_datetime('today') 
                    vuelos_pasados = hist_finca[hist_finca['FECHA_DT'] < fecha_vuelo_actual]
                    if not vuelos_pasados.empty:
                        ultima_fecha = vuelos_pasados['FECHA_DT'].max()
                        dias_ciclo_calc = (fecha_vuelo_actual - ultima_fecha).days

        # --- D. EXTRACCIÓN SEPARADA DE HECTÁREAS Y PISTA ---
        raw_pedido = str(datos_raw.get(20, datos_raw.get(21, "S/N"))).strip()
        num_pedido = raw_pedido.split('.')[0] 
        
        ha_cobro_detectada = float(datos_raw.get(8, 0)) # Hectáreas del INFORME DE PISTA (Col 8)
        ha_dosis_detectada = 0.0 # Nacerá del 459
        pista_detectada = str(datos_raw.get(2, "PLUC")).strip().upper() # Pista del informe
        lista_pistas_validas = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
        
        productos_pedido = pd.DataFrame()
        if not df_ped.empty and num_pedido != "S/N":
            filas_con_pedido = df_ped.astype(str).apply(lambda x: x.str.contains(num_pedido, case=False, na=False)).any(axis=1)
            productos_pedido = df_ped[filas_con_pedido]
            
            if not productos_pedido.empty:
                # 📡 Radar Pista (Busca las siglas en el pedido por si acaso)
                texto_total_pedido = productos_pedido.to_string().upper()
                for p_val in lista_pistas_validas:
                    if p_val in texto_total_pedido: pista_detectada = p_val; break
                        
                # 🎯 RADAR FRANCOTIRADOR CÓDIGO 459 (Hectáreas DOSIS/FACTURA)
                for _, fila_p in productos_pedido.iterrows():
                    col_material = [c for c in fila_p.index if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper()]
                    if col_material:
                        cod_item = str(fila_p[col_material[0]]).split('.')[0]
                        if "459" in cod_item:
                            col_cant = [c for c in fila_p.index if 'DOSIS' in str(c).upper() or 'CANT' in str(c).upper()]
                            if col_cant:
                                ha_dosis_detectada = extraer_numero(fila_p[col_cant[0]])
                                break # Encontrado
                
        # Si el 459 no estaba, usamos la de cobro como salvavidas temporal
        if ha_dosis_detectada == 0.0: ha_dosis_detectada = ha_cobro_detectada

        # --- 2. PANEL DE CONTROLES VISUAL ---
        with st.container(border=True):
            st.markdown("#### ⚙️ Controles de Vuelo y Parámetros")
            
            c_info1, c_info2, c_info3, c_info4 = st.columns(4)
            c_info1.info(f"🧑‍🌾 Productor: **{tipo_productor}**")
            # Días Ciclo (Editable por el comandante)
            dias_ciclo = c_info2.number_input("⏳ Días Ciclo:", value=int(dias_ciclo_calc), step=1)
            
            # Ha DOSIS (Viene del 459) - Rige la matriz y el multiplicador de facturación
            ha_dosis = c_info3.number_input("🧪 Ha (DOSIS/FACTURA):", value=float(ha_dosis_detectada), step=0.1, help="Extraída del CÓDIGO 459 del Pedido SAP")
            
            # Ha COBRO (Viene del Informe) - Rige el costo unitario del avión
            ha_cobro = c_info4.number_input("💰 Ha (COBRO/INFORME):", value=float(ha_cobro_detectada), step=0.1, help="Extraída del Informe de Pista para diluir la Orden de Servicio")
            
            c_ctrl1, c_ctrl2, c_ctrl3 = st.columns(3)
            lista_aviones = ["THRUS SR2", "PIPER PA 36-375", "CESSNA O PIPER PA", "AIR TRACTOR", "CESSNA ASA", "DRONE DATAROT", "DRONE GENESYS", "DRONE AVIL"]
            avion_sel = c_ctrl1.selectbox("✈️ Tipo de Avión:", lista_aviones)
            
            pista_sugerida = next((p for p in lista_pistas_validas if p in pista_detectada), "PLUC")
            pista_sel = c_ctrl2.selectbox("🛣️ Pista Operativa:", lista_pistas_validas, index=lista_pistas_validas.index(pista_sugerida))
            
            horometro = c_ctrl3.number_input("⏱️ Horómetro:", value=1.00, step=0.01)

        # --- 4. GRAN MATRIZ DE PRODUCTOS ---
        st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Dosis")
        
        if not productos_pedido.empty:
            idx_precio = -1; idx_lote = -1; idx_saldo = -1
            if not df_sab.empty:
                for j, col in enumerate(df_sab.columns):
                    col_str = str(col).upper()
                    if 'MAYOR' in col_str or 'PRECIO' in col_str: idx_precio = j
                    if 'LOTE' in col_str: idx_lote = j
                    if ('LIBRE' in col_str or 'SALDO' in col_str) and 'VALOR' not in col_str: idx_saldo = j

            sap_dict_pista = {}
            datos_extraidos_sap = []
            
            for _, fila_sap in productos_pedido.iterrows():
                col_material = [c for c in fila_sap.index if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper()]
                cod_item = str(fila_sap[col_material[0]]).split('.')[0] if col_material else str(fila_sap.iloc[1]).split('.')[0]
                if "459" in cod_item or "429" in cod_item: continue 
                
                col_cant = [c for c in fila_sap.index if 'DOSIS' in str(c).upper() or 'CANT' in str(c).upper()]
                cant_total = extraer_numero(fila_sap[col_cant[0]]) if col_cant else 0.0
                dosis_pista = cant_total / ha_dosis if ha_dosis > 0 else 0.0
                
                nombre_p = f"Item {cod_item}"
                if not df_sab.empty:
                    match_sabana = df_sab[df_sab.iloc[:, 0].astype(str).str.contains(cod_item, case=False, na=False)]
                    if match_sabana.empty: match_sabana = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(cod_item, case=False, na=False)).any(axis=1)]
                    if not match_sabana.empty:
                        col_nombre_sab = [c for c in match_sabana.columns if 'TEXTO' in str(c).upper() or 'DESC' in str(c).upper()]
                        if col_nombre_sab: nombre_p = str(match_sabana.iloc[0][col_nombre_sab[0]]).upper()
                
                nombre_limpio = nombre_p.split('*')[0].strip().replace(" ", "")
                sap_dict_pista[nombre_limpio] = dosis_pista
                datos_extraidos_sap.append({"cod": cod_item, "nombre": nombre_p, "nombre_limpio": nombre_limpio, "cant_total": cant_total})

            coctel_ganador = "SIN COINCIDENCIA"
            dosis_oficiales_coctel = {}
            claves_boro_zinc = ["BT", "BANATREL", "ZN", "ZINTRAC", "ZITRON"]
            tiene_acond_alto = any(any(clave in p for p in sap_dict_pista.keys()) for clave in claves_boro_zinc)
            
            if not df_mez.empty:
                dict_recetas = {}
                dict_lideres = {}
                for _, row in df_mez.iterrows():
                    if len(row) > 3:
                        cid = str(row.iloc[0]).strip().upper() 
                        p_tabla_clean = str(row.iloc[1]).strip().upper().replace(" ", "") 
                        d_tabla = extraer_numero(row.iloc[2]) 
                        es_lider = str(row.iloc[3]).strip().upper() == "X" 
                        if cid and cid != 'NAN':
                            if cid not in dict_recetas: dict_recetas[cid] = {}
                            dict_recetas[cid][p_tabla_clean] = d_tabla
                            if es_lider: dict_lideres[cid] = p_tabla_clean

                max_p = -999
                for iter_id, receta in dict_recetas.items():
                    es_valido = True; puntaje = 0; lider_db = dict_lideres.get(iter_id, "")
                    match_lider = False
                    if lider_db:
                        for k_sap in sap_dict_pista.keys():
                            if lider_db == k_sap or (len(k_sap)>=4 and lider_db in k_sap) or (len(lider_db)>=4 and k_sap in lider_db):
                                match_lider = True; break
                    if match_lider: puntaje += 1000
                    else: es_valido = False
                        
                    if es_valido:
                        for p_receta, d_esperada in receta.items():
                            if p_receta == "ACONDICIONADORSV": d_esperada = 0.06 if tiene_acond_alto else 0.02
                            elif p_receta == "ACEITEDICAM":
                                nums = re.findall(r'\d', iter_id)
                                if nums: d_esperada = float(nums[0])
                            elif p_receta == "IMBIOSILO": d_esperada = 1.5 if iter_id.startswith("IN") else 1.0
                            
                            match_receta = False; dose_matched = False
                            for k_sap, d_sap in sap_dict_pista.items():
                                if p_receta == k_sap or (len(k_sap)>=4 and p_receta in k_sap) or (len(p_receta)>=4 and k_sap in p_receta):
                                    match_receta = True
                                    if abs(d_sap - d_esperada) <= 0.2: dose_matched = True; break
                            if match_receta: puntaje += 50 if dose_matched else 10
                            else: es_valido = False; break
                    
                    if es_valido and puntaje > max_p:
                        max_p = puntaje; coctel_ganador = iter_id
                        dosis_oficiales_coctel = receta.copy()
                        for pr in dosis_oficiales_coctel:
                            if pr == "ACONDICIONADORSV": dosis_oficiales_coctel[pr] = 0.06 if tiene_acond_alto else 0.02
                            elif pr == "ACEITEDICAM":
                                nums = re.findall(r'\d', iter_id)
                                if nums: dosis_oficiales_coctel[pr] = float(nums[0])
                            elif pr == "IMBIOSILO": dosis_oficiales_coctel[pr] = 1.5 if iter_id.startswith("IN") else 1.0

            if coctel_ganador != "SIN COINCIDENCIA": st.success(f"🤖 **MOTOR IA:** Cóctel Ganador Detectado: **{coctel_ganador}**")
            else: st.warning("⚠️ **MOTOR IA:** No se encontró un Cóctel exacto. Buscando dosis estándar...")

            matriz_datos = []
            for item_data in datos_extraidos_sap:
                cod_item = item_data['cod']
                nombre_p = item_data['nombre']
                nombre_limpio = item_data['nombre_limpio']
                cant_total_pedido = item_data['cant_total']
                
                costo_unit = 0.0; lote_sap = "SIN LOTE EN PISTA"; saldo_sap = 0.0
                
                if not df_sab.empty:
                    match_sabana_global = df_sab[df_sab.iloc[:, 0].astype(str).str.contains(cod_item, case=False, na=False)]
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

                dosis_teorica = None
                if "FOSFO" in nombre_limpio and "ESTRES" in nombre_limpio: dosis_teorica = 1.0
                else:
                    for p_receta, d_oficial in dosis_oficiales_coctel.items():
                        if p_receta == nombre_limpio or (len(nombre_limpio)>=4 and p_receta in nombre_limpio) or (len(p_receta)>=4 and nombre_limpio in p_receta):
                            dosis_teorica = d_oficial; break
                    
                    if dosis_teorica is None and not df_mez.empty:
                        for _, row_m in df_mez.iterrows():
                            if len(row_m) > 6:
                                prod_gral = str(row_m.iloc[5]).strip().upper().replace(" ", "")
                                if prod_gral and prod_gral not in ['NAN', 'PRODUCTO2', '']:
                                    if prod_gral == nombre_limpio or (len(nombre_limpio)>=4 and prod_gral in nombre_limpio) or (len(prod_gral)>=4 and nombre_limpio in prod_gral):
                                        d_val = extraer_numero(row_m.iloc[6])
                                        if d_val > 0: dosis_teorica = d_val; break

                costo_margen = round(costo_unit * mult_material, 3)

                matriz_datos.append({
                    "A: Producto": nombre_p,
                    "B: Dosis/Ha (SAP)": round(dosis_teorica, 3) if dosis_teorica is not None else 0.0,
                    "C: X (Extra %)": 0.0,
                    "D: Dosis Total (Sistema)": 0.0, 
                    "E: Costo Unit (+Margen)": round(costo_margen, 3),
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
            # 🚀 LA MATRIZ DE QUÍMICOS SE CALCULA CON HA (DOSIS)
            df_matriz["D: Dosis Total (Sistema)"] = (df_matriz["B_Val"] * (1 + df_matriz["C_Val"]/100) * ha_dosis).round(3)
            
            # --- CÁLCULO SUMAPRODUCTO MEZCLA ---
            costo_mezcla_total = (df_matriz["D: Dosis Total (Sistema)"] * df_matriz["E: Costo Unit (+Margen)"]).sum()
            df_matriz = df_matriz.drop(columns=["B_Val", "C_Val"])

            edited_df = st.data_editor(
                df_matriz, key='editor_valid', 
                column_config={
                    "B: Dosis/Ha (SAP)": st.column_config.NumberColumn("Dosis/Ha", min_value=0.000, format="%.3f"),
                    "C: X (Extra %)": st.column_config.NumberColumn("Extra %", min_value=0.000, max_value=100.000, format="%.3f"),
                    "D: Dosis Total (Sistema)": st.column_config.NumberColumn("Dosis Ideal", format="%.3f"),
                    "E: Costo Unit (+Margen)": st.column_config.NumberColumn("Costo Unit (+Margen)", format="$ %.0f"),
                    "H: Saldo Real SAP": st.column_config.NumberColumn("Saldo SAP", format="%.3f"),
                    "I: Sugerido SAP (Total)": st.column_config.NumberColumn("Sugerido SAP (Total)", format="%.3f"),
                },
                disabled=["A: Producto", "D: Dosis Total (Sistema)", "E: Costo Unit (+Margen)", "G: Lotes (SAP)", "H: Saldo Real SAP", "I: Sugerido SAP (Total)"],
                use_container_width=True, hide_index=True
            )
            
            # ====================================================================
            # 💰 PANEL FINANCIERO FINAL
            # ====================================================================
            st.markdown("---")
            st.markdown("#### 💳 Resumen Financiero de la Misión")
            
            dict_precios = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA": 3036525, "AIR TRACTOR": 4665107, "CESSNA ASA": 3666600, "DRONE DATAROT": 84427, "DRONE GENESYS": 75518, "DRONE AVIL": 71280}
            precio_hora_base = dict_precios.get(avion_sel, 0)
            
            recargo_terrestre = 45000 if pista_sel in ["PDIV", "LUCI"] and "DRONE" not in avion_sel else 0
            
            # Fórmulas Económicas: EL VUELO POR HA SE DILUYE CON HA_COBRO
            tarifa_vuelo_bruta_ha = (precio_hora_base * horometro) / ha_cobro if ha_cobro > 0 else 0
            costo_vuelo_ha = (tarifa_vuelo_bruta_ha + recargo_terrestre) * mult_avion
            
            # LA FACTURA TOTAL SE MULTIPLICA CON HA_DOSIS
            costo_vuelo_total = costo_vuelo_ha * ha_dosis
            
            costo_serv_tec_ha = dias_ciclo * tarifa_serv_tec
            costo_serv_tec_total = costo_serv_tec_ha * ha_dosis
            
            total_factura = costo_mezcla_total + costo_vuelo_total + costo_serv_tec_total
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("🧪 Costo Total Mezcla", f"${costo_mezcla_total:,.0f}")
            m2.metric("✈️ Costo Vuelo / Ha (O.S.)", f"${costo_vuelo_ha:,.0f}", f"Avión: {mult_avion}x | Base: ${tarifa_vuelo_bruta_ha:,.0f}")
            m3.metric("👨‍🔬 Serv. Técnico / Ha", f"${costo_serv_tec_ha:,.0f}", f"Tarifa Base: ${tarifa_serv_tec:,.0f}")
            m4.metric("🔥 TOTAL FACTURA", f"${total_factura:,.0f}", f"Multiplicado por Ha DOSIS: {ha_dosis}")

        else:
            st.warning(f"🚨 No se encontraron productos en SAP para: {num_pedido}")

        st.markdown("---")
        if st.button("💾 DETONAR FACTURA Y GUARDAR HISTORIAL", type="primary", use_container_width=True):
            st.balloons()
            st.success(f"✅ ¡Operación Exitosa! Liquidación de {finca_sel} por un total de ${total_factura:,.0f} procesada.")
