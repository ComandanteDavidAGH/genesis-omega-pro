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

        # --- 📍 SELECTOR DE PISTA PARA DRONES Y AVIONES ---
        st.markdown("---")
        st.markdown("### 🛰️ Coordenadas de Lanzamiento Final")
        
        # 1. Definimos tipo_mision ANTES de usarlo en el st.info
        tipo_mision = "DRONE" if mision_solo_dron else "AVION"
        
        c_p1, c_p2 = st.columns(2)
        
        with c_p1:
            # Lista de pistas para elegir
            pistas_disponibles = ["PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2", "PROPIA"]
            pista_manual = st.selectbox("📍 Confirmar Pista de Operación:", pistas_disponibles, index=pistas_disponibles.index(pista_sel) if pista_sel in pistas_disponibles else 0)

        with c_p2:
            # Ahora sí, tipo_mision ya existe y no dará error
            st.info(f"🚀 Misión: {tipo_mision} | 📋 Referencia: {vuelo_ref}")

        if st.button("💾 DETONAR FACTURA Y GUARDAR EN BÓVEDA", type="primary", use_container_width=True):
            with st.spinner("🚀 Inyectando datos en TABLA 1 y APOYO simultáneamente..."):
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
                    hoja_maestra = boveda.worksheet("TABLA 1")

                    # 2. Preparación de Variables
                    fecha_str = fecha_operacion.strftime("%d/%m/%Y")
                    dia_sem = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"][fecha_operacion.weekday()]
                    num_sem = fecha_operacion.isocalendar()[1]
                    os_virtual = f"VIRT-{finca_limpia[:3]}-{datetime.now().strftime('%H%M')}"
                    
                    # 🔍 Radar de Finca (Bloque, Sector, Área Bruta)
                    bloque_f = ""; sector_f = ""; ha_bruta_f = ""
                    if not df_t2.empty:
                        match_f = df_t2[df_t2.iloc[:, 0].str.upper().str.strip() == finca_limpia.upper().strip()]
                        if not match_f.empty:
                            sector_f = match_f.iloc[0, 1]
                            ha_bruta_f = match_f.iloc[0, 2]
                            bloque_f = match_f.iloc[0, 3]

                    # 🧮 Cálculos de Rendimiento
                    ha_f = float(ha_dosis_final)
                    h_total_v = (ha_f / 10) if mision_solo_dron else 1.0
                    vol_total_gln = ha_f * 6
                    rend_min = h_total_v * 60
                    piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                    hk_f = "DR51" if "DATAROT" in tipo_mision else "DR52" if "GENESYS" in tipo_mision else "DR53" if "AVIL" in tipo_mision else "S/N"

                    # --- 🎯 MISIL 1: TABLA 1 (AZUL) ---
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

                    # --- 🎯 MISIL 2: TABLA DE APOYO (Fórmula Corregida) ---
                    fila_apoyo = [""] * 15
                    
                    # Esta fórmula detecta su propia fila y ajusta el número automáticamente
                    # Usamos punto y coma (;) y nombres en inglés para compatibilidad total
                    fila_apoyo[0] = "=IFERROR(ROW()-3; 0)" 
                    
                    fila_apoyo[1] = finca_limpia
                    fila_apoyo[2] = ha_f
                    fila_apoyo[3] = float(costo_por_ha)
                    fila_apoyo[5] = fecha_str
                    fila_apoyo[8] = coctel_ganador
                    fila_apoyo[10] = pista_manual
                    fila_apoyo[13] = tipo_mision
                    
                    # 4. EJECUCIÓN DEL IMPACTO
                    hoja_maestra.append_row(row_azul, value_input_option='USER_ENTERED')
                    hoja_apoyo.append_row(fila_apoyo, value_input_option='USER_ENTERED')

                    st.balloons()
                    st.success(f"✅ IMPACTO TOTAL CONFIRMADO. Referencia: {os_virtual}")
                    
                    if 'memoria_excel' in st.session_state:
                        del st.session_state['memoria_excel']

                except Exception as e_save:
                    st.error(f"🚨 Falla en el Gatillo de Guardado: {e_save}")
                    
import pandas as pd
import streamlit as st
import google.generativeai as genai
import json
from datetime import datetime
import re

# ==========================================
# --- INICIO DEL MÓDULO 3: GENESIS PRO ---
# ==========================================
st.divider()
st.header("🛰️ MÓDULO 3: SISTEMA GÉNESIS TOTAL")
st.subheader("Buzón de Recepción y Puesto de Control")

# 1. CARGA DE BASES DE DATOS CON ESCUDO DE MEMORIA
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
    
    # Esta hoja se necesita siempre para guardar datos
    hoja_maestra = boveda.worksheet("TABLA 1")
    
    # --- 🛡️ ESCUDO ANTI-BLOQUEOS (Error 429) ---
    if 'memoria_excel' not in st.session_state:
        with st.spinner("📡 Sincronizando bases de datos con la Bóveda (Solo una vez)..."):
            memoria = {}
            memoria['col_os'] = hoja_maestra.col_values(1)
            memoria['col_fincas'] = hoja_maestra.col_values(3)
            memoria['col_cocteles'] = hoja_maestra.col_values(7)
            
            try:
                d_t2 = boveda.worksheet("Tabla 2").get_all_values()
                memoria['df_t2'] = pd.DataFrame(d_t2[1:], columns=d_t2[0])
            except: memoria['df_t2'] = pd.DataFrame()

            try:
                d_t3 = boveda.worksheet("Tabla 3").get_all_values()
                memoria['df_t3'] = pd.DataFrame(d_t3[1:], columns=d_t3[0])
            except: memoria['df_t3'] = pd.DataFrame()

            try:
                d_apoyo = boveda.worksheet("TABLA DE APOYO2023").get_all_values()
                d_ap_limpio = [f + [""] * (15 - len(f)) if len(f) < 15 else f for f in d_apoyo]
                memoria['df_apoyo'] = pd.DataFrame(d_ap_limpio[9:])
            except: memoria['df_apoyo'] = pd.DataFrame()
            
            st.session_state['memoria_excel'] = memoria

    # Asignar variables desde la memoria
    mem = st.session_state['memoria_excel']
    lista_os_existentes = [str(os).strip() for os in mem['col_os'] if str(os).strip() != "" and str(os).upper() != "Nº ORDEN"]
    lista_fincas_oficiales = sorted(list(set([str(f).strip() for f in mem['col_fincas'] if str(f).strip() != "" and str(f).upper() != "FINCA"])))
    lista_cocteles_oficiales = sorted(list(set([str(c).strip() for c in mem['col_cocteles'] if str(c).strip() != "" and str(c).upper() != "COCTEL"])))
    df_t2 = mem['df_t2']
    df_t3 = mem['df_t3']
    df_apoyo = mem['df_apoyo']

except Exception as e:
    st.error(f"🚨 Falla de conexión principal: {e}")
    st.stop()

# 2. CEREBRO IA
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    modelo_ia = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    st.error("🚨 Falla en llaves IA.")
    st.stop()

# 3. RECEPCIÓN
archivo_os = st.file_uploader("📥 Subir Orden de Servicio", type=['pdf', 'jpg', 'jpeg', 'png'])

if archivo_os is not None:
    if st.button("🧠 ESCANEO DE INTELIGENCIA GÉNESIS", type="primary"):
        with st.spinner("🤖 Analizando documento..."):
            try:
                documento_bytes = archivo_os.getvalue()
                archivo_ia = [{"mime_type": archivo_os.type, "data": documento_bytes}]
                prompt = "Extrae datos de FUMIGARAY en JSON: fecha, numero_os, piloto, aeronave_hk, horometro_total, valor_hectarea, recargo, fincas:[{nombre_finca, hectareas}]."
                res = modelo_ia.generate_content([prompt, archivo_ia[0]], generation_config={"response_mime_type": "application/json"})
                st.session_state['datos_os_ia'] = json.loads(res.text)
                st.success("🎯 Lectura completada!")
            except Exception as e: st.error(f"Error IA: {e}")

# 4. PUESTO DE CONTROL
if 'datos_os_ia' in st.session_state:
    lista_ordenes = st.session_state['datos_os_ia']
    if isinstance(lista_ordenes, dict): lista_ordenes = [lista_ordenes]
    
    for i, datos in enumerate(lista_ordenes):
        with st.expander(f"📄 OS {datos.get('numero_os')}", expanded=True):
            col1, col2, col3 = st.columns(3)
            os_val = col1.text_input("Nº Orden", value=str(datos.get('numero_os', '')), key=f"os_{i}")
            fecha_val = col2.text_input("Fecha", value=str(datos.get('fecha', '')), key=f"fecha_{i}")
            piloto_val = col3.text_input("Piloto", value=str(datos.get('piloto', '')), key=f"piloto_{i}")
            
            col4, col5, col6 = st.columns(3)
            hk_val = col4.text_input("HK Aeronave", value=str(datos.get('aeronave_hk', '')), key=f"hk_{i}")
            horo_val = col5.text_input("Horómetro TOTAL", value=str(datos.get('horometro_total', '')), key=f"horo_{i}")
            costo_val = col6.text_input("Costo / Hectárea", value=str(datos.get('valor_hectarea', '')), key=f"costo_{i}")
            recargo_val = st.text_input("Recargo Dominical ($)", value=str(datos.get('recargo', '0')), key=f"recargo_{i}")

            # --- BOTÓN DE ENRIQUECIMIENTO TOTAL ---
            if st.button(f"🔍 ENRIQUECER DATOS (OS {os_val})", key=f"btn_enriquecer_{i}"):
                for f_item in datos.get('fincas', []):
                    nombre = str(f_item['nombre_finca']).strip().upper()
                    if not df_t2.empty:
                        match_t2 = df_t2[df_t2.iloc[:, 0].str.upper().str.strip() == nombre]
                        if not match_t2.empty:
                            f_item['bloque'] = match_t2.iloc[0, 3]
                            f_item['sector'] = match_t2.iloc[0, 1]
                            f_item['ha_bruta'] = match_t2.iloc[0, 2]
                            f_item['tipo_productor'] = match_t2.iloc[0, 5]
                    
                    if not df_apoyo.empty:
                        match_ap = df_apoyo[df_apoyo.iloc[:, 1].str.upper().str.strip() == nombre]
                        if not match_ap.empty: f_item['coctel'] = match_ap.iloc[-1, 8]
                
                st.session_state['datos_os_ia'] = lista_ordenes
                st.rerun()

            df_fincas = pd.DataFrame(datos.get('fincas', []))
            for c in ['bloque', 'sector', 'ha_bruta', 'coctel', 'tipo_productor']:
                if c not in df_fincas.columns: df_fincas[c] = ""

            df_editado = st.data_editor(
                df_fincas, use_container_width=True, num_rows="dynamic", key=f"ed_{i}",
                column_config={
                    "nombre_finca": st.column_config.SelectboxColumn("Finca", options=lista_fincas_oficiales),
                    "coctel": st.column_config.SelectboxColumn("Cóctel", options=lista_cocteles_oficiales)
                }
            )
            datos['fincas'] = df_editado.to_dict('records')

            # --- BOTÓN DE GUARDADO FINAL ---
            if str(os_val).strip() in lista_os_existentes:
                st.error("🚨 Esta OS ya existe en Excel.")
            else:
                if st.button(f"💾 GUARDAR TODO EN TABLA 1 (OS {os_val})", type="primary", key=f"save_{i}"):
                    try:
                        with st.spinner("🚀 Procesando cálculos y enviando a la Bóveda..."):
                            # 1. Limpieza de Fecha y Cálculos de Tiempo
                            f_raw = str(fecha_val).lower()
                            meses_dict = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06','julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}
                            fecha_corta = f_raw
                            for m_n, m_v in meses_dict.items():
                                if m_n in f_raw:
                                    nums = re.findall(r'\d+', f_raw)
                                    if len(nums) >= 2: fecha_corta = f"{nums[0].zfill(2)}/{m_v}/{nums[-1]}"
                                    break
                            
                            dt_obj = datetime.strptime(fecha_corta, "%d/%m/%Y")
                            dia_sem = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"][dt_obj.weekday()]
                            num_sem = dt_obj.isocalendar()[1]

                            # 2. Preparar Matemáticas
                            h_total = float(str(horo_val).replace(',','.'))
                            p_ha = float(str(costo_val).replace('.','').replace(',','.'))
                            rec_ha = float(str(recargo_val).replace('.','').replace(',','.'))
                            t_ha_os = sum([float(str(f['hectareas']).replace(',','.')) for f in datos['fincas']])

                            # 3. Radar de Búsqueda de Avión (En Tabla 2)
                            modelo_avion = "NO ENCONTRADO"
                            pista_avion = "NO ENCONTRADO"
                            hk_busqueda = str(hk_val).strip().upper()
                            
                            if not df_t2.empty:
                                # Col I:Matricula(8), J:Tipo(9), K:Pista(10)
                                match_avion = df_t2[df_t2.iloc[:, 8].str.upper().str.strip() == hk_busqueda]
                                if not match_avion.empty:
                                    modelo_avion = match_avion.iloc[0, 9]
                                    pista_avion = match_avion.iloc[0, 10]

                            filas = []
                            for f in datos['fincas']:
                                ha_f = float(str(f['hectareas']).replace(',','.'))
                                h_pro = (ha_f / t_ha_os) * h_total if t_ha_os > 0 else 0
                                
                                # --- 🧮 CÁLCULOS ESPECTACULARES ---
                                vol_total_gln = ha_f * 6
                                rend_minutos = h_pro * 60
                                costo_finca = (ha_f * p_ha) + (ha_f * rec_ha) # (F*T)+(F*U)
                                costo_avion_hora = costo_finca / h_pro if h_pro > 0 else 0 # V/N
                                costo_total_neto = p_ha * ha_f # T*F
                                
                                # --- 🗺️ PLANTILLA FINAL (A=0 hasta AH=33) ---
                                row = [""] * 34
                                row[0] = os_val               # A: Nº ORDEN
                                row[1] = f.get('bloque','')   # B: BLOQUE
                                row[2] = f['nombre_finca']    # C: FINCA
                                row[3] = f.get('sector','')   # D: SECTOR
                                row[4] = f.get('ha_bruta','') # E: ÀREA BRUTA
                                row[5] = ha_f                 # F: ÀREA FUMIGADA
                                row[6] = f.get('coctel','')   # G: COCTEL
                                row[7] = fecha_corta          # H: FECHA
                                row[8] = dia_sem              # I: DÌA SEM
                                row[9] = num_sem              # J: SEM
                                row[10] = h_total             # K: ODÒM. TOTAL
                                row[11] = 6                   # L: VOL. APLICADO gln/ha
                                row[12] = round(vol_total_gln, 2) # M: VOL. APLICADO gln
                                row[13] = round(h_pro, 2)     # N: RENDIMIENTO horas
                                row[14] = round(rend_minutos, 2) # O: RENDIMIENTO min
                                row[15] = piloto_val          # P: PILOTO
                                row[16] = hk_val              # Q: HK
                                row[17] = modelo_avion        # R: MODELO (Buscado por Python)
                                row[18] = pista_avion         # S: PISTA (Buscado por Python)
                                row[19] = p_ha                # T: $/ha
                                row[20] = rec_ha              # U: RECARGO $/ha
                                row[21] = round(costo_finca, 2) # V: COSTO AVIÒN / FINCA
                                row[22] = round(costo_avion_hora, 2) # W: COSTO AVIÒN / HORA
                                row[28] = round(costo_total_neto, 2)
                                # AA (26): VALOR A FACTURAR (Se deja vacío para el Modo Dios de Excel)
                                row[31] = 1                   # AF: Columna1
                                row[32] = f.get('tipo_productor','') # AG: TIPO DE PRODUCTOR
                                row[33] = "IA_GENESIS"        # AH: SISTEMA

                                filas.append(row)
                            
                            hoja_maestra.append_rows(filas, value_input_option='USER_ENTERED')
                            st.balloons()
                            st.success("✅ ¡Génesis ha completado el aterrizaje de datos!")
                            del st.session_state['memoria_excel']
                            
                    except Exception as e: st.error(f"Falla en guardado: {e}")
                        # =========================================================================
# --- 🔄 MÓDULO OMEGA V12: CLON EXACTO DE LA MACRO VBA ---
st.sidebar.markdown("---")
st.sidebar.subheader("📈 Sincronización Semanal")

semana_target = st.sidebar.select_slider("Semana a actualizar:", options=list(range(1, 53)), value=19)

def limpiar_texto_vba(t):
    """Clon exacto de la función LimpiarTexto de su Macro original"""
    if t is None: return ""
    # UCase y Trim
    temp = str(t).upper().strip()
    # Replace Chr(160) y puntos
    temp = temp.replace(chr(160), " ").replace(".", "")
    # Quitar dobles espacios
    while "  " in temp:
        temp = temp.replace("  ", " ")
    return temp

def val_seguro(v):
    """Convierte de forma segura a número sin dañar los decimales"""
    try: return float(v)
    except: return 0.0

if st.sidebar.button("🚀 EJECUTAR OMEGA V12", use_container_width=True):
    try:
        with st.spinner(f"Sincronizando Semana {semana_target} al estilo Macro..."):
            # 1. ORIGEN: Génesis (Pidiendo los números PUROS, sin formato de texto)
            url_gen = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
            sh_gen = gc.open_by_url(url_gen)
            
            # --- Precios (Configuración) ---
            raw_config = sh_gen.worksheet("Configuración").get_all_values(value_render_option='UNFORMATTED_VALUE')
            dict_precios = {}
            for row in raw_config:
                if len(row) > 9:
                    prod = limpiar_texto_vba(row[8]) # Columna I
                    if prod and prod != "PRODUCTO":
                        dict_precios[prod] = val_seguro(row[9]) # Columna J

            # --- Dosis (DD_Mesclas desde fila 13) ---
            raw_mezclas = sh_gen.worksheet("DD_Mesclas").get_all_values(value_render_option='UNFORMATTED_VALUE')
            dict_dosis = {}
            for row in raw_mezclas[12:]: # Empieza en Fila 13 real
                if len(row) > 10:
                    prod_m = limpiar_texto_vba(row[9]) # Columna J
                    if prod_m:
                        dict_dosis[prod_m] = val_seguro(row[10]) # Columna K

            # 2. DESTINO: Bóveda de Comparación
            url_dest = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
            sh_dest = gc.open_by_url(url_dest)
            ws_datos = sh_dest.worksheet("DATOS")
            
            # Pedimos los datos puros para no alterar la estructura
            datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
            
            # Buscar semana en Fila 7
            col_semana = -1
            for i, v in enumerate(datos_dest[6]):
                if str(v).strip() == str(semana_target):
                    col_semana = i + 1
                    break
            
            if col_semana == -1:
                st.error(f"❌ No se halló la semana {semana_target} en la Fila 7.")
            else:
                updates = []
                # 3. LÓGICA DE DETECCIÓN EXACTA (Igual a la Macro)
                for r_idx, row in enumerate(datos_dest):
                    n_fila = r_idx + 1
                    if n_fila < 8 or len(row) < 4: continue
                    
                    tipo_tabla = limpiar_texto_vba(row[1]) # Columna B (POR LITRO o DOSIS-HA)
                    producto_dest = limpiar_texto_vba(row[3]) # Columna D
                    
                    if not producto_dest: continue
                    
                    if producto_dest in dict_precios:
                        precio_unitario = dict_precios[producto_dest]
                        
                        # Si en la Columna B dice DOSIS-HA, multiplicamos.
                        if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                            # Estamos en la Tabla 2
                            if producto_dest in dict_dosis:
                                dosis_valor = dict_dosis[producto_dest]
                                valor_final = precio_unitario * dosis_valor
                            else:
                                valor_final = precio_unitario
                        else:
                            # Estamos en la Tabla 1
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

 # --- ✈️ MÓDULO OMEGA V20: INYECCIÓN EXACTA POR COORDENADAS ---
import datetime
import re

st.sidebar.markdown("---")
st.sidebar.subheader("✈️ Rastreo Dominicales")

url_ori = st.sidebar.text_input(
    "🔗 Pegue URL de GÉNESIS_OMEGA_V2_ESTABLE:", 
    placeholder="Pegue aquí el link y presione ENTER..."
)

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
            f = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))
            return f if f.year > 2020 else None
            
        v_str = str(v).lower().strip()
        if v_str.replace('.', '').isdigit():
            f = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(float(v_str)))
            return f if f.year > 2020 else None
        
        meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
        for mes, num_mes in meses.items():
            if mes in v_str:
                match_ano = re.search(r'\d{4}', v_str)
                match_dia = re.search(r'\b\d{1,2}\b', v_str)
                if match_ano and match_dia:
                    f = datetime.datetime(int(match_ano.group()), num_mes, int(match_dia.group()))
                    return f if f.year > 2020 else None
                    
        if "/" in v_str or "-" in v_str:
            import dateutil.parser
            f = dateutil.parser.parse(v_str, dayfirst=True)
            return f if f.year > 2020 else None
    except: pass
    return None

if st.sidebar.button("🚀 RASTREAR FALTANTES", use_container_width=True):
    if not url_ori or "http" not in url_ori:
        st.sidebar.error("❌ Pegue una URL válida y presione ENTER.")
    else:
        try:
            with st.spinner("Modo Inyección Exacta Activado..."):
                url_dest = "https://docs.google.com/spreadsheets/d/1FTiKlHo2UF8lWHk4SrFf9oxTUa2Q_n1l5IK9XFoqQaU/edit"
                
                # --- 1. LEER DESTINO ---
                sh_dest = gc.open_by_url(url_dest)
                ws_dest = sh_dest.sheet1
                datos_dest = ws_dest.get_all_values(value_render_option='UNFORMATTED_VALUE')
                
                max_f = datetime.datetime(1900, 1, 1)
                dict_local = {}
                
                for i, row in enumerate(datos_dest):
                    row_padded = row + [""] * (5 - len(row)) if len(row) < 5 else row
                    if i + 1 >= 5 and str(row_padded[1]).strip() != "":
                        f_obj = procesar_fecha_pesada(row_padded[3])
                        if f_obj:
                            if f_obj > max_f: max_f = f_obj
                            dict_local[f"{str(row_padded[1]).strip().upper()}|{f_obj.date()}"] = i + 1

                st.info(f"📅 Radar Destino: Última fecha validada -> {max_f.strftime('%d/%m/%Y')}")

                # --- 2. LEER ORIGEN CON RAYOS X ---
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

                # --- 3. INYECCIÓN TÁCTICA EXACTA ---
                if dict_nuevos:
                    # Calculamos la próxima fila vacía basada en la longitud real de los datos
                    prox_fila = len(datos_dest) + 1 
                    
                    # Quitamos el espacio vacío inicial. Ahora la lista empieza directamente con 'finca'
                    filas_nuevas = [[v['finca'], v['ha'], v['fec'], v['sur'], v['pista'], v['semana']] for v in dict_nuevos.values()]
                    
                    # Disparamos el comando update indicando que empiece exactamente en la Columna B
                    ws_dest.update(f'B{prox_fila}', filas_nuevas, value_input_option='USER_ENTERED')
                    
                    st.success(f"🎯 ¡IMPACTO PERFECTO! {len(filas_nuevas)} registros inyectados empezando en la fila {prox_fila}.")
                    st.balloons()
                else:
                    st.warning("⚠️ El escáner vio los recargos, pero ninguno era posterior a la fecha del radar.")

        except Exception as e:
            st.error(f"🚨 FALLA DE SISTEMA: {type(e).__name__} - {str(e)}")

# --- ⚖️ MÓDULO OMEGA: ARQUEO DE INVENTARIOS V14 (EL AUDITOR INTEGRAL) ---
import pandas as pd
import streamlit as st
import io
import unicodedata
import re

# 🕶️ Gafas de Visión Nocturna y Diseño de Tablas
st.markdown(
    """
    <style>
    div[data-baseweb="input"] input, div[data-baseweb="select"] {
        color: black !important;
        background-color: white !important;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.sidebar.markdown("---")
st.sidebar.subheader("⚖️ Arqueo de Inventarios V14")

archivo_sap = st.sidebar.file_uploader("1️⃣ Sábana de SAP", type=['xlsx', 'csv'])
archivos_sup = st.sidebar.file_uploader("2️⃣ Reportes Supervisores (.xlsx)", type=['xlsx'], accept_multiple_files=True)
semana_obj = st.sidebar.text_input("🎯 Semana a Auditar (Ej: 17):", placeholder="Escriba aquí...")

# --- SISTEMA DE MEMORIA BLINDADA ---
if "arqueo_procesado" not in st.session_state:
    st.session_state.arqueo_procesado = False

def purificar_lote(lote):
    if pd.isna(lote) or lote is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(lote).upper().strip())

def quitar_tildes(s):
    if pd.isna(s) or s is None: return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(s).upper().strip()) if unicodedata.category(c) != 'Mn')

# --- MOTOR MATEMÁTICO DE CRUCE ---
def generar_cruce():
    cruce = pd.merge(st.session_state.df_sap_grouped, st.session_state.df_sup_grouped, on=['PISTA', 'LOTE_KEY'], how='outer')
    
    cruce['ITEM'] = cruce['ITEM'].fillna("---")
    cruce['PRODUCTO'] = cruce['PRODUCTO'].fillna(cruce['PRODUCTO_SUP']).fillna("N/A")
    cruce['LOTE'] = cruce['LOTE'].fillna(cruce['LOTE_SUP'])
    cruce['SALDO_SAP'] = cruce['SALDO_SAP'].fillna(0).round(2)
    cruce['SALDO_FISICO'] = cruce['SALDO_FISICO'].fillna(0).round(2)
    cruce['DIFERENCIA'] = (cruce['SALDO_FISICO'] - cruce['SALDO_SAP']).round(2)
    
    # 📝 NUEVA COLUMNA DE OBSERVACIONES
    if 'OBSERVACIONES' not in cruce.columns:
        cruce['OBSERVACIONES'] = ""
        
    cruce['ESTADO'] = cruce['DIFERENCIA'].apply(lambda x: "✅ OK" if abs(x) <= 0.05 else "❌ DISCREPANCIA")
    cruce = cruce[['PISTA', 'ITEM', 'PRODUCTO', 'LOTE_KEY', 'LOTE', 'SALDO_SAP', 'SALDO_FISICO', 'DIFERENCIA', 'ESTADO', 'OBSERVACIONES']]
    cruce = cruce.sort_values(by=['PISTA', 'PRODUCTO'])
    
    # Si ya existían observaciones previas guardadas, las respetamos
    if "observaciones_memoria" in st.session_state:
        # Combinar para no borrar lo que el usuario ya escribió
        for idx, row in cruce.iterrows():
            key = f"{row['PISTA']}_{row['LOTE_KEY']}"
            if key in st.session_state.observaciones_memoria:
                cruce.at[idx, 'OBSERVACIONES'] = st.session_state.observaciones_memoria[key]

    st.session_state.cruce_final = cruce

# --- BOTÓN DE INICIO RADAR ---
if st.sidebar.button("🚀 INICIAR ARQUEO ESTRATÉGICO", use_container_width=True):
    if not archivo_sap or not archivos_sup or not semana_obj:
        st.sidebar.error("❌ Faltan suministros (Archivos o Semana).")
    else:
        try:
            with st.spinner("Escaneando y agrupando bases de datos..."):
                # Limpiar memoria de observaciones anteriores si es un nuevo cruce
                st.session_state.observaciones_memoria = {}
                
                # --- FASE 1: SAP ---
                sap_file = archivo_sap[0] if isinstance(archivo_sap, list) else archivo_sap
                df_sap = pd.read_csv(sap_file) if sap_file.name.endswith('.csv') else pd.read_excel(sap_file)
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

                # --- FASE 2: SUPERVISORES ---
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
                            row = [quitar_tildes(x) for x in df_raw.iloc[i].values if pd.notna(x)]
                            if any("LOTE" in val for val in row) and any("SALDO" in val for val in row):
                                h_idx = i; break
                        
                        if h_idx != -1:
                            df_s = df_raw.iloc[h_idx + 1:].copy()
                            df_s.columns = [f"{quitar_tildes(x)}_{idx}" for idx, x in enumerate(df_raw.iloc[h_idx])]
                            
                            c_p = next((c for c in df_s.columns if "PRODUC" in c or "DESCRI" in c), None)
                            c_a = next((c for c in df_s.columns if "ALMAC" in c or "PISTA" in c), None)
                            c_l = next((c for c in df_s.columns if "LOTE" in c and "SALDO" not in c), None)
                            c_v = next((c for c in df_s.columns if "SALDO" in c and "INIC" not in c and "SAP" not in c), None)
                            
                            if all([c_p, c_a, c_l, c_v]):
                                df_s_c = df_s[[c_p, c_a, c_l, c_v]].copy()
                                df_s_c.columns = ['PRODUCTO_SUP', 'PISTA', 'LOTE_SUP', 'SALDO_FISICO']
                                
                                df_s_c['PISTA'] = df_s_c['PISTA'].astype(str).str.strip().str.upper().replace('NAN', None).replace('NONE', None).replace('', None).ffill().bfill()
                                df_s_c['LOTE_KEY'] = df_s_c['LOTE_SUP'].apply(purificar_lote)
                                df_s_c['SALDO_FISICO'] = pd.to_numeric(df_s_c['SALDO_FISICO'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                                
                                lista_sup.append(df_s_c)

                if not lista_sup:
                    st.error("🚨 No se encontraron datos válidos.")
                else:
                    df_sup_total = pd.concat(lista_sup, ignore_index=True)
                    st.session_state.df_sup_grouped = df_sup_total.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                    
                    st.session_state.semana_actual = semana_obj
                    generar_cruce()
                    st.session_state.arqueo_procesado = True

        except Exception as e:
            st.error(f"🚨 FALLA EN EL SISTEMA: {type(e).__name__} - {e}")

# --- FASE 3 Y 4: PANEL INTERACTIVO DE FUSIÓN Y OBSERVACIONES ---
if st.session_state.arqueo_procesado:
    st.success(f"🎯 Auditoría Semana {st.session_state.semana_actual} Cargada. Proceda al Flujo de Trabajo.")
    
    # 💡 FLUJO SUGERIDO PARA EL COMANDANTE
    st.info("📌 **FLUJO SUGERIDO:** 1️⃣ Primero vaya a la pestaña 'Conciliador' y fusione los lotes mal escritos. 2️⃣ Luego vuelva a 'Reporte' y escriba sus observaciones.")
    
    tab1, tab2 = st.tabs(["📊 Reporte Oficial y Observaciones", "🛠️ Conciliador de Lotes Fantasmas"])
    
    with tab1:
        st.subheader("Diferencias Reales (Escriba sus observaciones directo en la tabla)")
        
        df_mostrar = st.session_state.cruce_final[st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA"].drop(columns=['LOTE_KEY'])
        
        if df_mostrar.empty:
            st.balloons()
            st.success("✅ ¡TODO EL INVENTARIO ESTÁ CUADRADO! No hay diferencias.")
            df_final_descarga = st.session_state.cruce_final.drop(columns=['LOTE_KEY'])
        else:
            # 📝 TABLA EDITABLE SOLO PARA OBSERVACIONES
            edited_df = st.data_editor(
                df_mostrar,
                use_container_width=True,
                hide_index=True,
                disabled=["PISTA", "ITEM", "PRODUCTO", "LOTE", "SALDO_SAP", "SALDO_FISICO", "DIFERENCIA", "ESTADO"],
                column_config={
                    "OBSERVACIONES": st.column_config.TextColumn(
                        "📝 OBSERVACIONES",
                        help="Haga clic aquí para escribir el motivo de la diferencia.",
                        max_chars=200
                    )
                }
            )
            
            # Guardamos las observaciones en memoria en tiempo real
            if "observaciones_memoria" not in st.session_state:
                st.session_state.observaciones_memoria = {}
                
            # Almacenamos lo que escribe el usuario usando PISTA + LOTE como llave
            # Como borramos LOTE_KEY para mostrar, reconstruimos la llave con purificar_lote
            for idx, row in edited_df.iterrows():
                key = f"{row['PISTA']}_{purificar_lote(row['LOTE'])}"
                st.session_state.observaciones_memoria[key] = row['OBSERVACIONES']
                
            # Aplicamos las observaciones al dataframe final en la memoria
            for idx, row in st.session_state.cruce_final.iterrows():
                key = f"{row['PISTA']}_{row['LOTE_KEY']}"
                if key in st.session_state.observaciones_memoria:
                    st.session_state.cruce_final.at[idx, 'OBSERVACIONES'] = st.session_state.observaciones_memoria[key]

            df_final_descarga = st.session_state.cruce_final.drop(columns=['LOTE_KEY'])
        
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final_descarga[df_final_descarga['ESTADO'] == "❌ DISCREPANCIA"].to_excel(writer, index=False, sheet_name='Diferencias_Activas')
            df_final_descarga.to_excel(writer, index=False, sheet_name='Arqueo_Total')
        st.download_button("📥 Descargar Reporte Final (Con Observaciones)", buffer.getvalue(), f"Arqueo_Semana_{st.session_state.semana_actual}.xlsx")

    with tab2:
        st.markdown("### 👻 Lotes Fantasmas (Causados por errores de digitación)")
        
        # 🎯 FILTRO TÁCTICO: Solo mostramos donde SAP tiene 0 pero el físico tiene saldo
        errores_fantasmas = st.session_state.cruce_final[
            (st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA") & 
            (st.session_state.cruce_final['SALDO_SAP'] == 0) & 
            (st.session_state.cruce_final['SALDO_FISICO'] > 0)
        ]
        
        if errores_fantasmas.empty:
            st.success("✅ No hay 'Lotes Fantasmas'. Las diferencias que quedan en el Reporte Oficial son descuadres reales de inventario.")
        else:
            opciones_error = errores_fantasmas.apply(
                lambda x: f"Pista: {x['PISTA']} | Prod: {x['PRODUCTO']} | Lote Fantasma: {x['LOTE']} | Qty: {x['SALDO_FISICO']}", axis=1
            ).tolist()
            
            error_sel = st.selectbox("1️⃣ Seleccione el lote mal escrito por el supervisor:", opciones_error)
            
            if error_sel:
                idx = opciones_error.index(error_sel)
                row_sel = errores_fantasmas.iloc[idx]
                
                pista_target = row_sel['PISTA']
                prod_target = row_sel['PRODUCTO']
                lote_key_target = row_sel['LOTE_KEY']
                
                lotes_sap_filtrados = st.session_state.df_sap_raw[
                    (st.session_state.df_sap_raw['PISTA'] == pista_target) & 
                    (st.session_state.df_sap_raw['PRODUCTO'] == prod_target)
                ]['LOTE'].dropna().unique().tolist()
                
                if len(lotes_sap_filtrados) > 0:
                    lote_correcto = st.selectbox(f"2️⃣ Seleccione el Lote Oficial de SAP para fusionar:", lotes_sap_filtrados)
                    
                    if st.button("⚡ FUSIONAR Y DESTRUIR FANTASMA", type="primary"):
                        mask = (st.session_state.df_sup_grouped['PISTA'] == pista_target) & \
                               (st.session_state.df_sup_grouped['LOTE_KEY'] == lote_key_target)
                               
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_SUP'] = lote_correcto
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_KEY'] = purificar_lote(lote_correcto)
                        
                        st.session_state.df_sup_grouped = st.session_state.df_sup_grouped.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        
                        generar_cruce()
                        st.rerun()
                else:
                    st.warning("⚠️ SAP no envió este producto a esta pista. El supervisor lo inventó o lo movió sin avisar. (Debe justificarlo en Observaciones).")
