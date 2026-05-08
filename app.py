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
                    st.session_state['df_sabana'] = pd.read_excel(bytes_sabana) if f_sabana.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(bytes_sabana, sep=None, engine='python')
                    
                    # 2. Leer Pedidos
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos) if f_pedidos.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    # 3. CONEXIÓN SATELITAL
                    try:
                        if "gcp_credentials" in st.secrets:
                            cred_dict = dict(st.secrets["gcp_credentials"])
                            gc = gspread.service_account_from_dict(cred_dict)
                        else:
                            gc = gspread.service_account(filename='credenciales.json')
                        
                        url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                        boveda = gc.open_by_url(url_boveda)
                        
                        # Cargar Configuración (TABLA 2)
                        hoja_tabla2 = boveda.worksheet("TABLA 2")
                        datos_tabla2 = hoja_tabla2.get_all_values()
                        st.session_state['df_config'] = pd.DataFrame(datos_tabla2[1:], columns=datos_tabla2[0])
                        
                        # 🔥 ESCÁNER INTELIGENTE DE TÍTULOS (TABLA DE APOYO) 🔥
                        hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023") 
                        datos_apoyo = hoja_apoyo.get_all_values()
                        
                        # 1. Buscar en qué fila están realmente los títulos
                        fila_titulos = 0
                        for i, fila in enumerate(datos_apoyo[:20]): # Escanea las primeras 20 filas
                            if any('FINCA' in str(celda).upper() for celda in fila):
                                fila_titulos = i
                                break
                                
                        encabezados_crudos = datos_apoyo[fila_titulos]
                        
                        # 2. Escudo Anti-Duplicados y Columnas Vacías
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
                                
                        # 3. Crear la tabla cortando la basura de arriba
                        st.session_state['df_apoyo'] = pd.DataFrame(datos_apoyo[fila_titulos+1:], columns=encabezados_limpios)
                        
                        # Cargar Mezclas
                        hoja_mezclas = boveda.worksheet("DD_Mesclas")
                        datos_mezclas = hoja_mezclas.get_all_values()
                        st.session_state['df_mezclas'] = pd.DataFrame(datos_mezclas[1:], columns=datos_mezclas[0])
                        
                        # Cargar Configuración Base
                        hoja_conf = boveda.worksheet("Configuración")
                        datos_conf = hoja_conf.get_all_values()
                        st.session_state['df_config_base'] = pd.DataFrame(datos_conf[1:], columns=datos_conf[0])
                        
                        st.success("🛰️ Enlace Satelital Establecido con Escáner de Títulos.")
                        
                    except Exception as error_nube:
                        st.error(f"🚨 Falla en el Enlace Satelital: {error_nube}")
                        
                    # 4. ESCÁNER DE PISTAS (Sigue igual porque funciona)
                    lista_pistas = []
                    for f in f_pistas:
                        dict_p = pd.read_excel(io.BytesIO(f.getvalue()), sheet_name=None, header=None)
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
        # 🟢 MOTOR DE INTELIGENCIA Y LÓGICA
        # =========================================================================
        import re
        from datetime import datetime
        
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

        def parse_fecha_pesada(val):
            if pd.isna(val) or str(val).strip() == "": return pd.NaT
            if isinstance(val, (datetime, pd.Timestamp)): return pd.to_datetime(val)
            s = str(val).lower().strip()
            if s.isnumeric(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), unit='D')
            meses = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}
            s_clean = s.replace(',', '').replace('del', '').replace('de', '')
            mes_encontrado = next((num for mes, num in meses.items() if mes in s_clean), None)
            if mes_encontrado:
                nums = re.findall(r'\d+', s_clean)
                anio = next((n for n in nums if len(n) == 4), None)
                dia = next((n for n in nums if len(n) <= 2), None)
                if anio and dia: return pd.to_datetime(f"{anio}-{mes_encontrado}-{dia.zfill(2)}")
            try: return pd.to_datetime(s, dayfirst=True)
            except: return pd.NaT

        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        finca_limpia = re.sub(r'\s+', ' ', str(finca_sel)).strip().upper()

        # --- A. PRODUCTOR Y TOPE ---
        tipo_productor = "REVISAR FINCA"
        tipo_de_tope_finca = "SIN TOPE"
        if not df_t2.empty:
            match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia]
            if not match_t2.empty:
                fila_t2 = match_t2.iloc[0]
                tipo_productor = str(fila_t2.iloc[5]).strip().upper()
                tipo_de_tope_finca = str(fila_t2.iloc[6]).strip().upper()

        # --- B. MULTIPLICADORES ---
        mult_material = 1.112; tarifa_serv_tec_base = 1337.0; mult_avion_base = 1.112
        if not df_cfg.empty:
            match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
            if not match_cfg.empty:
                fila_c = match_cfg.iloc[0]
                mult_material = extraer_numero(fila_c.iloc[3])
                tarifa_serv_tec_base = extraer_numero(fila_c.iloc[4])
                mult_avion_base = extraer_numero(fila_c.iloc[6])

        # --- C. CAZADOR DE DÍAS CICLO ---
        dias_ciclo_calc = 0
        if not df_apoyo.empty:
            col_finca = [c for c in df_apoyo.columns if 'FINCA' in str(c).upper()]
            col_fecha = [c for c in df_apoyo.columns if 'FECHA' in str(c).upper()]
            if col_finca and col_fecha:
                mask_finca = df_apoyo[col_finca[0]].apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip().upper()) == finca_limpia
                hist_finca = df_apoyo[mask_finca].copy()
                if not hist_finca.empty:
                    hist_finca['FECHA_DT'] = hist_finca[col_fecha[0]].apply(parse_fecha_pesada)
                    hist_finca = hist_finca.dropna(subset=['FECHA_DT'])
                    if not hist_finca.empty:
                        fecha_ref = pd.to_datetime(fecha_operacion)
                        vuelos_anteriores = hist_finca[hist_finca['FECHA_DT'] < fecha_ref]
                        if not vuelos_anteriores.empty:
                            dias_ciclo_calc = (fecha_ref - vuelos_anteriores['FECHA_DT'].max()).days

        # --- D. EXTRACCIÓN DE DATOS PISTA ---
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
                texto_pedido = match_ped.to_string().upper()
                for p_val in lista_pistas_validas:
                    if p_val in texto_pedido: pista_detectada = p_val; break
                for _, r_p in match_ped.iterrows():
                    if len(r_p) >= 7 and "459" in str(r_p.iloc[5]):
                        ha_dosis_detectada = extraer_numero(r_p.iloc[6])
                        break
        
        ha_cobro_detectada = extraer_numero(datos_raw.get(8, 0))
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        # =========================================================================
        # 🟢 2. PANELES DE CONTROL LOGÍSTICO
        # =========================================================================
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
            
            # --- INTERRUPTOR MULTI-AVIÓN ---
            multi_aviones = r1c4.toggle("✈️ Recargo Coord. Multi-Avión", value=False, key=f"ma_{casilla_key}")
            mult_avion_final = mult_avion_base + 0.1 if multi_aviones else mult_avion_base

            # --- PARÁMETROS TERRESTRES (SE OCULTAN SI ES 100% DRON) ---
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

        # --- DICCIONARIOS DE TARIFAS ---
        dict_topes_pista = {"TOPE MAX GENERAL": {"PLUC": 63325, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325}, "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517}, "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}}
        val_tope = dict_topes_pista.get(tipo_de_tope_finca, {}).get(pista_sel, 999999)
        dict_aviones = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA": 3036525, "AIR TRACTOR": 4665107, "CESSNA ASA": 3666600}
        dict_drones = {"DRONE DATAROT": 84427, "DRONE GENESYS": 75518, "DRONE AVIL": 71280}

        # --- 2.5 HANGAR DINÁMICO (ADAPTABLE) ---
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
                    # El dron SÍ toma el multiplicador de la finca/productor
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
                
                # Liquidar Aviones
                for _, row in escuadron_aviones.iterrows():
                    av_sel, ha_av, horo = row["Avión"], float(row.get("Hectáreas", 0)), float(row.get("Horómetro", 0))
                    if pd.isna(av_sel) or ha_av <= 0: continue
                    total_ha_cobro_escuadron += ha_av
                    tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av
                    tarifa_aplicada = tarifa_base_ha + recargo_final if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope) + recargo_final
                    costo_total_vuelos += (tarifa_aplicada * ha_av) * mult_avion_final
                
                # Liquidar Drones de Apoyo
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    costo_total_vuelos += (dict_drones.get(dr_sel, 0) * ha_dr) * mult_avion_final

        # =========================================================================
        # 🟢 3. MATRIZ DE MEZCLA
        # =========================================================================
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

                costo_margen = costo_unit * mult_material

                matriz_datos.append({
                    "A: Producto": nombre_p,
                    "B: Dosis/Ha (SAP)": round(dosis_teorica, 3) if dosis_teorica is not None else 0.0,
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

            # --- RAMPA DE EXTRACCIÓN AL PORTAPAPELES ---
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("##### 📋 Copia Rápida para SAP (Costo Unitario)")
            costos_limpios = df_matriz['E: Costo Unit (+Margen)'].fillna(0).astype(int).astype(str).tolist()
            texto_para_copiar = "\n".join(costos_limpios)
            st.code(texto_para_copiar, language="text")
            st.caption("☝️ Pase el mouse por la esquina superior derecha del cuadro oscuro arriba y haga clic en el icono de copiar (📋). Luego vaya a SAP y presione Ctrl+V.")

        else:
            st.warning("🚨 No se encontró un pedido válido para la matriz de químicos.")
            costo_mezcla_total = 0.0

        # =========================================================================
        # 🟢 4. LIQUIDACIÓN FINAL Y MÉTRICAS
        # =========================================================================
        st.markdown("---")
        st.markdown("### 💰 Liquidación Final (Bóveda SAP)")
        
        tarifa_st_final = d_ciclo_factura * tarifa_serv_tec_base
        subtotal_st = tarifa_st_final * ha_dosis_final
        gran_total = costo_mezcla_total + costo_total_vuelos + subtotal_st
        costo_por_ha = gran_total / ha_dosis_final if ha_dosis_final > 0 else 0

        # --- PANELES DE MÉTRICA ---
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("🚜 Hectáreas Cobro Totales", f"{total_ha_cobro_escuadron:.2f} Ha")
        
        if mision_solo_dron:
            r2.metric("🛣️ Condición Pista", "NO APLICA (Dron)")
        else:
            r2.metric("🛣️ Condición Pista", tipo_de_tope_finca, f"Límite: $ {fmt_sap(val_tope)}")
            
        r3.metric("👨‍🔬 Tarifa Serv. Tec (Base)", f"$ {fmt_sap(tarifa_serv_tec_base)}")
        r4.metric("✈️ Multiplicador Aplicado", f"x {mult_avion_final}")

        st.markdown("<br>", unsafe_allow_html=True)
        c_sap1, c_sap2, c_sap3, c_sap4 = st.columns(4)
        
        with c_sap1:
            st.caption("🧪 Mezcla Total")
            st.code(fmt_sap(costo_mezcla_total), language=None)
        with c_sap2:
            st.caption("✈️ Costo Total de Vuelo")
            st.code(fmt_sap(costo_total_vuelos), language=None)
        with c_sap3:
            st.caption("👨‍🔬 Costo Serv. Técnico")
            st.code(fmt_sap(subtotal_st), language=None)
        with c_sap4:
            st.markdown(f"""
            <div style='background-color:#0d1b2a; padding:10px; border-radius:5px; border:1px solid #d4af37; text-align:center;'>
                <p style='margin:0; color:#d4af37; font-size:12px;'>💰 COSTO x HECTÁREA</p>
                <h4 style='margin:0; color:white;'>$ {fmt_sap(costo_por_ha)}</h4>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.metric("🔥 TOTAL FACTURACIÓN FINCA (GRAN TOTAL)", f"$ {fmt_sap(gran_total)}", f"Basado en {ha_dosis_final} Ha")

        if st.button("💾 DETONAR FACTURA Y GUARDAR EN BÓVEDA", type="primary", use_container_width=True):
            with st.spinner("🚀 Inyectando datos de costos en la Bóveda Satelital (Google Drive)..."):
                try:
                    # 1. Reconexión Satelital
                    if "gcp_credentials" in st.secrets:
                        cred_dict = dict(st.secrets["gcp_credentials"])
                        gc = gspread.service_account_from_dict(cred_dict)
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                    
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023")

                    # 2. Identificar el Tipo de Misión
                    if mision_solo_dron:
                        tipo_mision = "DRONE"
                        # Nota: Si es solo dron, pista_sel por defecto es "PLUC" u otra que haya quedado guardada.
                        # Si necesita que diga algo específico como "N/A" para drones, me avisa.
                    else:
                        tipo_mision = "AVION"

                    # 3. Armar el Misil de Datos (CORRECCIÓN COLUMNA K = PISTA)
                    # Columnas: A, B:Finca, C:Ha, D, E:Costo, F:Fecha, G, H, I:Coctel, J, K:PISTA, L, M, N:Tipo
                    fila_apoyo = [
                        "",                                     # Col 1 (A)
                        finca_limpia,                           # Col 2 (B): Finca
                        float(ha_dosis_final),                  # Col 3 (C): Hectáreas
                        "",                                     # Col 4 (D)
                        float(gran_total),                      # Col 5 (E): Costo Total
                        fecha_operacion.strftime("%d/%m/%Y"),   # Col 6 (F): Fecha
                        "", "",                                 # Col 7 (G), 8 (H)
                        coctel_ganador,                         # Col 9 (I): Coctel
                        "",                                     # Col 10 (J)
                        pista_sel,                              # Col 11 (K): 🎯 NOMBRE DE LA PISTA
                        "", "",                                 # Col 12 (L), 13 (M)
                        tipo_mision                             # Col 14 (N): "DRONE" o "AVION"
                    ]

                    # 4. Disparo a la nube
                    hoja_apoyo.append_row(fila_apoyo)

                    st.balloons()
                    st.success(f"✅ ¡MISIÓN GUARDADA! El costo de la finca {finca_limpia} (Pista: {pista_sel}) ha sido inyectado en la TABLA DE APOYO2023.")
                    
                except Exception as e_save:
                    st.error(f"🚨 Falla en el Gatillo de Guardado: {e_save}")
import pandas as pd
import streamlit as st
import google.generativeai as genai
import json

# --- INICIO DEL MÓDULO 3 ---
st.divider()
st.header("🛰️ MÓDULO 3: RADAR DE ÓRDENES DE SERVICIO (Visión IA)")
st.subheader("Buzón de Recepción y Puesto de Control")

# 1. CARGA DE BASE DE DATOS (FRANCOTIRADOR TRIPLE)
try:
    if "gcp_credentials" in st.secrets:
        cred_dict = dict(st.secrets["gcp_credentials"])
        import gspread
        gc = gspread.service_account_from_dict(cred_dict)
    else:
        import gspread
        gc = gspread.service_account(filename='credenciales.json')
    
    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    boveda = gc.open_by_url(url_boveda)
    hoja_maestra = boveda.worksheet("TABLA 1")
    
    columna_os = hoja_maestra.col_values(1)
    columna_fincas = hoja_maestra.col_values(3)
    columna_cocteles = hoja_maestra.col_values(7)
    
    lista_os_existentes = [str(os).strip() for os in columna_os if str(os).strip() != "" and str(os).upper() != "Nº ORDEN"]
    
    lista_cocteles_oficiales = []
    for c in columna_cocteles:
        c_limpio = str(c).strip()
        if c_limpio != "" and c_limpio.upper() != "COCTEL" and c_limpio not in lista_cocteles_oficiales:
            lista_cocteles_oficiales.append(c_limpio)
            
except Exception as e:
    st.error(f"🚨 Falla de conexión con la Bóveda Satelital: {e}")
    lista_os_existentes = []
    lista_cocteles_oficiales = []

# 2. CONFIGURACIÓN DEL CEREBRO IA
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    modelo_ia = genai.GenerativeModel('gemini-2.5-pro')
except Exception as e:
    st.error("🚨 Falla en el sistema de IA. Revise sus llaves de seguridad.")
    st.stop()

# 3. BUZÓN DE RECEPCIÓN (Dropzone)
archivo_os = st.file_uploader("📥 Arrastre aquí la foto o PDF de la Orden de Servicio", type=['pdf', 'jpg', 'jpeg', 'png'])

if archivo_os is not None:
    st.success("✅ Documento recibido en la bahía de carga.")
    
    if st.button("🚀 LANZAR DRON DE RECONOCIMIENTO (Señuelo)", type="primary"):
        with st.spinner("🤖 El Dron está sobrevolando el documento y tomando notas. Por favor espere..."):
            try:
                documento_bytes = archivo_os.getvalue()
                tipo_mime = archivo_os.type
                archivo_ia = [{"mime_type": tipo_mime, "data": documento_bytes}]
                
                # 📜 EL SEÑUELO: Sin reglas estrictas, solo que nos diga qué ve con sus propios ojos.
                orden_militar = """
                Analiza esta imagen como si fueras un humano leyendo el papel. Respóndeme SOLO estas 4 preguntas de forma clara y directa:
                1. ¿Qué fecha exacta ves en la parte de arriba del todo? (Escribe la frase completa que leas, letra por letra).
                2. Busca la sección '3- INFORMACION FUMIGACION'. ¿Qué número está escrito EXACTAMENTE DEBAJO del título 'Rendimiento Hectareas/Hora'?
                3. ¿Ves algún número cerca de la palabra 'Recargo' en la parte de abajo?
                4. ¿Cuál es el número del Horómetro Total?
                """
                
                # Disparamos el señuelo sin formato JSON para que hable libremente
                respuesta = modelo_ia.generate_content([orden_militar, archivo_ia[0]])
                
                st.warning("🚨 REPORTE EN CRUDO DEL DRON (Lo que la IA ve realmente):")
                st.info(respuesta.text)
                
            except Exception as e:
                st.error(f"❌ El Dron fue derribado por la interferencia: {e}")

# 4. EL PUESTO DE CONTROL (Escaneo Múltiple con Recargo)
if 'datos_os_ia' in st.session_state:
    datos_ia = st.session_state['datos_os_ia']
    
    if isinstance(datos_ia, dict):
        lista_ordenes = [datos_ia]
    elif isinstance(datos_ia, list):
        lista_ordenes = datos_ia
    else:
        lista_ordenes = []
        
    st.write("### 🚦 PUESTO DE CONTROL: Verifique los datos extraídos")
    st.info(f"📡 El radar detectó **{len(lista_ordenes)}** Órdenes de Servicio en este documento.")
    
    for i, datos in enumerate(lista_ordenes):
        st.markdown(f"#### 📄 Documento #{i+1}: Orden de Servicio {datos.get('numero_os', 'Desconocida')}")
        
        # Fila 1
        col1, col2, col3 = st.columns(3)
        os_leida = col1.text_input("Nº Orden", value=datos.get('numero_os', ''), key=f"os_{i}")
        fecha_leida = col2.text_input("Fecha", value=datos.get('fecha', ''), key=f"fecha_{i}")
        col3.text_input("Piloto", value=datos.get('piloto', ''), key=f"piloto_{i}")
        
        # Fila 2
        col4, col5, col6 = st.columns(3)
        col4.text_input("HK Aeronave", value=datos.get('aeronave_hk', ''), key=f"hk_{i}")
        col5.text_input("Horómetro TOTAL", value=datos.get('horometro_total', ''), help="Solo la diferencia", key=f"horo_{i}")
        col6.text_input("Costo / Hectárea", value=datos.get('valor_hectarea', ''), key=f"costo_{i}")
        
        # Fila 3 (El Recargo)
        col7, col8, col9 = st.columns(3)
        col7.text_input("Recargo ($)", value=datos.get('recargo', ''), help="Dejar en 0 si no hay", key=f"recargo_{i}")
        
        st.write(f"**Fincas de la OS {os_leida}:**")
        df_fincas = pd.DataFrame(datos.get('fincas', []))
        
        df_fincas_editado = st.data_editor(
            df_fincas, 
            use_container_width=True, 
            num_rows="dynamic",
            key=f"tabla_fincas_{i}",
            column_config={
                "coctel": st.column_config.SelectboxColumn(
                    "Cóctel (Menú Oficial)",
                    help="Elija el cóctel correcto",
                    options=lista_cocteles_oficiales,
                    required=True
                )
            }
        )
        
        # 🛡️ ESCUDO ANTI-DUPLICADOS
        os_limpia = str(os_leida).strip()
        if os_limpia in lista_os_existentes:
            st.error(f"🚨 ¡ALERTA! La OS Nº '{os_limpia}' ya existe en su Excel. No se puede duplicar.")
        else:
            st.success("✅ OS autorizada.")
            if st.button(f"🚀 APROBAR OS {os_limpia} Y PREPARAR PRORRATEO", type="primary", key=f"btn_aprobar_{i}"):
                st.info("¡Motor listo para disparar los datos a la TABLA 1!")
                
        st.divider()
