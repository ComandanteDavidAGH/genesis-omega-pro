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
    
    if 'df_pistas' not in st.session_state or 'df_apoyo' not in st.session_state:
        st.warning("🚨 Cargue los archivos en el Módulo 1 e inicie el procesamiento.")
    else:
        # --- 1. PANEL DE SELECCIÓN ---
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

        # =========================================================================
        # 🟢 MOTOR DE INTELIGENCIA DINÁMICA (CON REGLA DE ORO)
        # =========================================================================
        import re
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

        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        finca_limpia = str(finca_sel).strip().upper()

        # --- A. PRODUCTOR Y TOPE ---
        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        if not df_t2.empty:
            match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).str.strip().str.upper() == finca_limpia]
            if not match_t2.empty:
                fila_t2 = match_t2.iloc[0]
                tipo_productor = str(fila_t2.iloc[5]).strip().upper()
                tipo_de_tope_finca = str(fila_t2.iloc[6]).strip().upper()

        # --- B. MULTIPLICADORES ---
        mult_material = 1.112; tarifa_serv_tec_base = 1337.0; mult_avion = 1.112
        if not df_cfg.empty:
            match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
            if not match_cfg.empty:
                fila_c = match_cfg.iloc[0]
                mult_material = extraer_numero(fila_c.iloc[3])
                tarifa_serv_tec_base = extraer_numero(fila_c.iloc[4])
                mult_avion = extraer_numero(fila_c.iloc[6])

        # --- C. DÍAS CICLO (MOTOR DE TABLA DE APOYO 2023) ---
        dias_ciclo_calc = 0
        if not df_apoyo.empty:
            col_finca = [c for c in df_apoyo.columns if 'FINCA' in str(c).upper()]
            col_fecha = [c for c in df_apoyo.columns if 'FECHA' in str(c).upper()]
            if col_finca and col_fecha:
                hist_finca = df_apoyo[df_apoyo[col_finca[0]].astype(str).str.strip().str.upper() == finca_limpia].copy()
                if not hist_finca.empty:
                    hist_finca['FECHA_DT'] = pd.to_datetime(hist_finca[col_fecha[0]], errors='coerce')
                    hist_finca = hist_finca.dropna(subset=['FECHA_DT'])
                    if not hist_finca.empty:
                        fecha_ref = pd.to_datetime(fecha_operacion)
                        vuelos_anteriores = hist_finca[hist_finca['FECHA_DT'] < fecha_ref]
                        if not vuelos_anteriores.empty:
                            dias_ciclo_calc = (fecha_ref - vuelos_anteriores['FECHA_DT'].max()).days

        # --- D. HECTÁREAS 459 Y PISTA (PEDIDOS SAP) ---
        datos_vuelo = vuelos_informe[vuelos_informe['ORIGEN'] == vuelo_ref].iloc[0]
        datos_raw = datos_vuelo['DATOS_FILA']
        num_pedido = str(datos_raw.get(20, datos_raw.get(21, "S/N"))).split('.')[0]
        
        lista_pistas_validas = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
        pista_detectada = "PLUC"
        ha_dosis_detectada = 0.0
        match_ped = pd.DataFrame()

        if not df_ped.empty and num_pedido != "S/N":
            match_ped = df_ped[df_ped.astype(str).apply(lambda x: x.str.contains(num_pedido)).any(axis=1)]
            if not match_ped.empty:
                # Pista
                texto_pedido = match_ped.to_string().upper()
                for p_val in lista_pistas_validas:
                    if p_val in texto_pedido: pista_detectada = p_val; break
                # Hectáreas 459 (Col F y G)
                for _, r_p in match_ped.iterrows():
                    # Buscamos en la Columna F (Material / Índice 5)
                    val_material = str(r_p.iloc[5]).strip()
                    if "459" in val_material:
                        # Extraemos de la Columna G (Cantidad Pendiente / Índice 6)
                        ha_dosis_detectada = extraer_numero(r_p.iloc[6])
                        break

        ha_cobro_detectada = extraer_numero(datos_raw.get(8, 0))
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        # --- 2. PANEL CONTROLES (CASILLAS DINÁMICAS POR KEY) ---
        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros de Operación e Inteligencia de Ciclos")
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            r1c1.info(f"🧑‍🌾 Productor: **{tipo_productor}**")
            r1c2.warning(f"🛣️ Tope Finca: **{tipo_de_tope_finca}**")
            
            # El uso de 'key' con variables obliga a refrescar el valor al cambiar de finca
            casilla_key = f"{finca_sel}_{vuelo_ref}_{fecha_operacion}"
            
            r1c3.number_input("📅 Ciclo (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key=f"ds_{casilla_key}")
            d_ciclo_factura = r1c4.number_input("⏳ Ciclo (COBRO)", value=int(dias_ciclo_calc), step=1, key=f"df_{casilla_key}")

            r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns(5)
            lista_aviones = ["THRUS SR2", "PIPER PA 36-375", "CESSNA O PIPER PA", "AIR TRACTOR", "CESSNA ASA", "DRONE DATAROT", "DRONE GENESYS", "DRONE AVIL"]
            avion_sel = r2c1.selectbox("✈️ Avión", lista_aviones, key=f"av_{casilla_key}")
            
            pista_sugerida = next((p for p in lista_pistas_validas if p in pista_detectada), "PLUC")
            pista_sel = r2c2.selectbox("🛣️ Pista", lista_pistas_validas, index=lista_pistas_validas.index(pista_sugerida), key=f"pi_{casilla_key}")
            
            horometro = r2c3.number_input("⏱️ Horómetro", value=1.00, key=f"hor_{casilla_key}")
            ha_dosis_final = r2c4.number_input("🧪 Ha Dosis (459)", value=float(ha_dosis_detectada), key=f"had_{casilla_key}")
            ha_cobro_final = r2c5.number_input("💰 Ha Cobro (Inf)", value=float(ha_cobro_detectada), key=f"hac_{casilla_key}")

            st.markdown("##### 🚛 Porción Terrestre / Recargo")
            rec_col1, rec_col2 = st.columns([1, 1])
            idx_recargo = 1 if pista_sel == "PDIV" else 0 
            opciones_rec = ["0 (Sin Recargo)", "8504 (Porción PDIV)", "45000 (Recargo T. General)", "Otro Valor Manual..."]
            recargo_lista = rec_col1.selectbox("Seleccione Valor:", opciones_rec, index=idx_recargo, key=f"rl_{casilla_key}")
            if recargo_lista == "Otro Valor Manual...":
                recargo_final = rec_col2.number_input("✍️ Digite Recargo ($)", value=0, step=1000, key=f"rm_{casilla_key}")
            else:
                recargo_final = float(recargo_lista.split(" ")[0])

        # --- 3. MATRIZ DE MEZCLA ---
        st.markdown("#### 🧪 Matriz de Mezcla")
        # Aquí continúa su lógica de matriz... (costo_mezcla_total calculado con ha_dosis_final)
        costo_mezcla_total = 0.0 # Se asume cálculo previo

        # --- 4. TOPES (REGLA DE ORO: PRECIOS EXACTOS) ---
        dict_topes_pista = {
            "TOPE MAX GENERAL": {"PLUC": 63325, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325},
            "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517},
            "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}
        }
        val_tope = dict_topes_pista.get(tipo_de_tope_finca, {}).get(pista_sel, 999999)

        st.markdown("---")
        st.markdown("### 💰 Liquidación Final (Bóveda SAP)")
        
        dict_precios = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA": 3036525, "AIR TRACTOR": 4665107, "CESSNA ASA": 3666600, "DRONE DATAROT": 84427, "DRONE GENESYS": 75518, "DRONE AVIL": 71280}
        p_hora = dict_precios.get(avion_sel, 0)

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("⏱️ Precio Base Avión (Hora)", f"$ {fmt_sap(p_hora)}")
        limite_display = f"Límite: $ {fmt_sap(val_tope)}" if val_tope != 999999 else "Sin Límite"
        r2.metric("🛣️ Condición Pista", tipo_de_tope_finca, limite_display)
        r3.metric("🚛 Recargo Terrestre", f"$ {fmt_sap(recargo_final)}")
        r4.metric("👨‍🔬 Tarifa Serv. Tec (Base)", f"$ {fmt_sap(tarifa_serv_tec_base)}")

        tarifa_vuelo_ha = (p_hora * horometro) / ha_cobro_final if ha_cobro_final > 0 else 0
        if pista_sel == "PDIV": tarifa_final_vuelo = (tarifa_vuelo_ha + recargo_final) * mult_avion
        else: tarifa_final_vuelo = (min(tarifa_vuelo_ha, val_tope) + recargo_final) * mult_avion

        tarifa_st_final = d_ciclo_factura * tarifa_serv_tec_base

        c_sap1, c_sap2, c_sap3 = st.columns(3)
        with c_sap1:
            st.caption("🧪 Mezcla Total")
            st.code(fmt_sap(costo_mezcla_total), language=None)
        with c_sap2:
            st.caption("✈️ Vuelo x Ha (O.S.)")
            st.code(fmt_sap(tarifa_final_vuelo), language=None)
        with c_sap3:
            st.caption("👨‍🔬 Serv. Técnico x Ha")
            st.code(fmt_sap(tarifa_st_final), language=None)

        gran_total = costo_mezcla_total + (tarifa_final_vuelo * ha_dosis_final) + (tarifa_st_final * ha_dosis_final)
        st.metric("🔥 TOTAL A FACTURAR FINCA", f"$ {fmt_sap(gran_total)}", f"Calculado sobre {ha_dosis_final} Ha")
