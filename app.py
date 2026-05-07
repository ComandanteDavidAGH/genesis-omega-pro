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
    
    # --- RADAR DE DIAGNÓSTICO (NUEVO BLINDAJE) ---
    faltan_datos = False
    if 'df_pistas' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró el Informe de Pistas en la memoria. ¿Presionó el botón de procesar en el Módulo 1?")
        faltan_datos = True
    if 'df_apoyo' not in st.session_state:
        st.error("🚨 ALERTA: No se encontró la 'TABLA DE APOYO2023' de Google Drive. Revise la conexión satelital en el Módulo 1.")
        faltan_datos = True
        
    if faltan_datos:
        st.warning("⚠️ Vuelva al Módulo 1, cargue los archivos, presione INICIAR PROCESAMIENTO y espere el mensaje verde de éxito antes de volver aquí.")
    else:
        # --- 1. UBICACIÓN ESTRATÉGICA DE DATOS GENERALES ---
        with st.container(border=True):
            st.markdown("### 📡 Panel de Operaciones (Datos de Vuelo)")
            c1, c2 = st.columns(2)
            
            # LISTA DEPLEGABLE DE FINCAS (Desde TABLA DE APOYO2023 / TABLA DE APOYO2023)
            lista_fincas_apoyo = st.session_state['df_apoyo'].iloc[:, 1].unique().tolist() # Asumiendo columna 0 es Finca
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

# --- 4. GRAN MATRIZ DE PRODUCTOS (MOTOR MACRO HIJO) ---
            st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Dosis")

            raw_pedido = str(datos_raw.get(20, datos_raw.get(21, "S/N"))).strip()
            num_pedido = raw_pedido.split('.')[0] 
            
            df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
            df_sab = st.session_state.get('df_sabana', pd.DataFrame())
            df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
            df_cfg = st.session_state.get('df_config_base', pd.DataFrame()) 
            
            def extraer_numero(valor):
                if pd.isna(valor) or valor == "": return 0.0
                if isinstance(valor, (int, float)): return float(valor)
                v = str(valor).strip().upper()
                if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
                elif ',' in v: v = v.replace(',', '.') 
                try: return float(v)
                except: return 0.0

            tipo_productor = "SOCIO"
            if 'df_apoyo' in st.session_state and not st.session_state['df_apoyo'].empty:
                df_apoyo = st.session_state['df_apoyo']
                match_finca = df_apoyo[df_apoyo.iloc[:, 1].astype(str).str.contains(str(finca_sel), case=False, na=False)]
                if not match_finca.empty:
                    col_tipo = [c for c in df_apoyo.columns if 'TIPO' in str(c).upper() or 'GRUPO' in str(c).upper()]
                    if col_tipo: tipo_productor = str(match_finca.iloc[0][col_tipo[0]])

            productos_pedido = pd.DataFrame()
            if not df_ped.empty and num_pedido != "S/N":
                filas_con_pedido = df_ped.astype(str).apply(lambda x: x.str.contains(num_pedido, case=False, na=False)).any(axis=1)
                productos_pedido = df_ped[filas_con_pedido]

            if not productos_pedido.empty:
                # ====================================================================
                # 🧠 FASE 1: PRE-ESCANEO
                # ====================================================================
                sap_dict_pista = {}
                datos_extraidos_sap = []
                
                for _, fila_sap in productos_pedido.iterrows():
                    col_material = [c for c in fila_sap.index if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper() or 'CÓDIGO' in str(c).upper()]
                    cod_item = str(fila_sap[col_material[0]]).split('.')[0] if col_material else str(fila_sap.iloc[1]).split('.')[0]
                    
                    if "459" in cod_item or "429" in cod_item: continue 
                    
                    col_cant = [c for c in fila_sap.index if 'DOSIS' in str(c).upper() or 'CANT' in str(c).upper()]
                    cant_total = extraer_numero(fila_sap[col_cant[0]]) if col_cant else 0.0
                    dosis_pista = cant_total / ha_real if ha_real > 0 else 0.0
                    
                    nombre_p = f"Item {cod_item}"
                    if not df_sab.empty:
                        match_sabana = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(cod_item, case=False, na=False)).any(axis=1)]
                        if not match_sabana.empty:
                            col_nombre_sab = [c for c in match_sabana.columns if 'TEXTO' in str(c).upper() or 'DESC' in str(c).upper()]
                            if col_nombre_sab: nombre_p = str(match_sabana.iloc[0][col_nombre_sab[0]]).upper()
                    
                    nombre_limpio = nombre_p.split('*')[0].strip().replace(" ", "")
                    sap_dict_pista[nombre_limpio] = dosis_pista
                    
                    datos_extraidos_sap.append({
                        "cod": cod_item, "nombre": nombre_p, "nombre_limpio": nombre_limpio, "cant_total": cant_total
                    })

                # ====================================================================
                # ⚔️ FASE 2: LA GRAN BATALLA (Scoring System)
                # ====================================================================
                coctel_ganador = "SIN COINCIDENCIA"
                dosis_oficiales_coctel = {}
                
                # 🚨 NUEVA REGLA DE ORO EXTENDIDA (BT, ZN, ZINTRAC, ZITRON, BANATREL)
                claves_boro_zinc = ["BT", "BANATREL", "ZN", "ZINTRAC", "ZITRON"]
                tiene_acond_alto = any(any(clave in p for clave in claves_boro_zinc) for p in sap_dict_pista.keys())
                
                if not df_mez.empty:
                    dict_recetas = {}
                    dict_lideres = {}
                    
                    for _, row in df_mez.iterrows():
                        cid = str(row.iloc[0]).strip().upper() 
                        if not cid or cid == 'NAN': continue
                        p_tabla_clean = str(row.iloc[5]).strip().upper().replace(" ", "") 
                        d_tabla = extraer_numero(row.iloc[6]) 
                        es_lider = str(row.iloc[3]).strip().upper() == "X" 
                        
                        if cid not in dict_recetas: dict_recetas[cid] = {}
                        dict_recetas[cid][p_tabla_clean] = d_tabla
                        if es_lider: dict_lideres[cid] = p_tabla_clean

                    max_p = -999
                    import re
                    
                    for iter_id, receta in dict_recetas.items():
                        es_valido = True
                        puntaje = 0
                        lider_db = dict_lideres.get(iter_id, "")
                        
                        match_lider = False
                        if lider_db:
                            for k_sap in sap_dict_pista.keys():
                                if lider_db == k_sap or (len(k_sap)>=4 and lider_db in k_sap) or (len(lider_db)>=4 and k_sap in lider_db):
                                    match_lider = True
                                    break
                        if match_lider: puntaje += 1000
                        else: es_valido = False
                            
                        if es_valido:
                            for p_receta, d_esperada in receta.items():
                                # APLICACIÓN DE REGLAS DE ORO
                                if p_receta == "ACONDICIONADORSV": d_esperada = 0.06 if tiene_acond_alto else 0.02
                                elif p_receta == "ACEITEDICAM":
                                    nums = re.findall(r'\d+', iter_id)
                                    if nums: d_esperada = float(nums[0])
                                elif p_receta == "IMBIOSILO": d_esperada = 1.5 if iter_id.startswith("IN") else 1.0
                                
                                match_receta = False
                                dose_matched = False
                                for k_sap, d_sap in sap_dict_pista.items():
                                    if p_receta == k_sap or (len(k_sap)>=4 and p_receta in k_sap) or (len(p_receta)>=4 and k_sap in p_receta):
                                        match_receta = True
                                        if abs(d_sap - d_esperada) <= 0.2: dose_matched = True
                                        break
                                
                                if match_receta: puntaje += 50 if dose_matched else 10
                                else:
                                    es_valido = False
                                    break
                        
                        if es_valido and puntaje > max_p:
                            max_p = puntaje
                            coctel_ganador = iter_id
                            dosis_oficiales_coctel = receta.copy()
                            for pr in dosis_oficiales_coctel:
                                if pr == "ACONDICIONADORSV": dosis_oficiales_coctel[pr] = 0.06 if tiene_acond_alto else 0.02
                                elif pr == "ACEITEDICAM":
                                    nums = re.findall(r'\d+', iter_id)
                                    if nums: dosis_oficiales_coctel[pr] = float(nums[0])
                                elif pr == "IMBIOSILO": dosis_oficiales_coctel[pr] = 1.5 if iter_id.startswith("IN") else 1.0

                if coctel_ganador != "SIN COINCIDENCIA":
                    st.success(f"🤖 **MOTOR IA:** Cóctel Ganador Detectado: **{coctel_ganador}**")
                else:
                    st.warning("⚠️ **MOTOR IA:** No se encontró un Cóctel exacto. Se usarán dosis referenciales.")

                # ====================================================================
                # 🏗️ FASE 3: CONSTRUCCIÓN DE MATRIZ FINAL
                # ====================================================================
                matriz_datos = []
                for item_data in datos_extraidos_sap:
                    cod_item = item_data['cod']
                    nombre_p = item_data['nombre']
                    nombre_limpio = item_data['nombre_limpio']
                    cant_total_pedido = item_data['cant_total']
                    
                    costo_unit = 0.0
                    lote_sap = "S/L"
                    saldo_sap = 0.0
                    
                    if not df_sab.empty:
                        match_sabana = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(cod_item, case=False, na=False)).any(axis=1)]
                        if not match_sabana.empty:
                            fila_sabana = match_sabana.iloc[0]
                            col_precio = [c for c in fila_sabana.index if str(c).strip().upper() in ['PRECIOS', 'PRECIO']]
                            if col_precio: costo_unit = extraer_numero(fila_sabana[col_precio[0]])
                            col_lote = [c for c in fila_sabana.index if 'LOTE' in str(c).upper()]
                            if col_lote: lote_sap = str(fila_sabana[col_lote[0]])
                            col_saldo = [c for c in fila_sabana.index if 'LIBRE' in str(c).upper() or 'SALDO' in str(c).upper()]
                            if col_saldo: saldo_sap = extraer_numero(fila_sabana[col_saldo[0]])

                    dosis_teorica = None
                    for p_receta, d_oficial in dosis_oficiales_coctel.items():
                        if p_receta == nombre_limpio or (len(nombre_limpio)>=4 and p_receta in nombre_limpio) or (len(p_receta)>=4 and nombre_limpio in p_receta):
                            dosis_teorica = d_oficial
                            break
                    
                    multiplicador_margen = 1.112
                    if not df_cfg.empty:
                        match_prod = df_cfg[df_cfg[0].astype(str).str.contains(tipo_productor, case=False, na=False)]
                        if not match_prod.empty:
                            multiplicador_margen = extraer_numero(match_prod.iloc[0][3])
                            if multiplicador_margen == 0.0: multiplicador_margen = 1.112
                    costo_margen = round(costo_unit * multiplicador_margen, 3)

                    matriz_datos.append({
                        "A: Producto": nombre_p,
                        "B: Dosis/Ha (SAP)": round(dosis_teorica, 3) if dosis_teorica is not None else None,
                        "C: X (Extra %)": None,
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
                
                df_matriz["D: Dosis Total (Sistema)"] = (df_matriz["B_Val"] * (1 + df_matriz["C_Val"]/100) * ha_real).round(3)
                df_matriz = df_matriz.drop(columns=["B_Val", "C_Val"])

                edited_df = st.data_editor(
                    df_matriz,
                    key='editor_valid', 
                    column_config={
                        "B: Dosis/Ha (SAP)": st.column_config.NumberColumn("Dosis/Ha", min_value=0.000, format="%.3f"),
                        "C: X (Extra %)": st.column_config.NumberColumn("Extra %", min_value=0.000, max_value=100.000, step=0.001, format="%.3f"),
                        "D: Dosis Total (Sistema)": st.column_config.NumberColumn("Dosis Ideal", format="%.3f"),
                        "E: Costo Unit (+Margen)": st.column_config.NumberColumn("Costo Unit (+Margen)", format="$ %.0f"),
                        "H: Saldo Real SAP": st.column_config.NumberColumn("Saldo SAP", format="%.3f"),
                        "I: Sugerido SAP (Total)": st.column_config.NumberColumn("Sugerido SAP (Total)", format="%.3f"),
                    },
                    disabled=["A: Producto", "D: Dosis Total (Sistema)", "E: Costo Unit (+Margen)", "G: Lotes (SAP)", "H: Saldo Real SAP", "I: Sugerido SAP (Total)"],
                    use_container_width=True, hide_index=True
                )
            else:
                st.warning(f"🚨 No se encontraron productos en SAP para: {num_pedido}")

            st.markdown("---")
            if st.button("🔥 DETONAR FACTURA Y GUARDAR HISTORIAL", type="primary", use_container_width=True):
                st.balloons()
                st.success(f"✅ ¡Operación Exitosa! Liquidación de {finca_sel} procesada.")
