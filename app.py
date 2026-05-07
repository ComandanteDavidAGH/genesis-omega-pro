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
        # 🟢 MOTOR DE INTELIGENCIA (EXTRACTOR BLINDADO)
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

        # 🔥 TRADUCTOR DEFINITIVO (Fuerza Bruta para cualquier formato)
        def parse_fecha_pesada(val):
            if pd.isna(val) or str(val).strip() == "": return pd.NaT
            if isinstance(val, (datetime, pd.Timestamp)): return pd.to_datetime(val)
            s = str(val).lower().strip()
            
            if s.isnumeric(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), unit='D')
            
            meses = {'enero':'01','febrero':'02','marzo':'03','abril':'04','mayo':'05','junio':'06',
                     'julio':'07','agosto':'08','septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'}
            
            # Limpiar basura textual
            s_clean = s.replace(',', '').replace('del', '').replace('de', '')
            
            mes_encontrado = None
            for mes, num in meses.items():
                if mes in s_clean:
                    mes_encontrado = num
                    break
                    
            if mes_encontrado:
                nums = re.findall(r'\d+', s_clean)
                if len(nums) >= 2:
                    anio = next((n for n in nums if len(n) == 4), None)
                    dia = next((n for n in nums if len(n) <= 2), None)
                    if anio and dia:
                        return pd.to_datetime(f"{anio}-{mes_encontrado}-{dia.zfill(2)}")
                        
            try: return pd.to_datetime(s, dayfirst=True)
            except: return pd.NaT

        df_ped = st.session_state.get('df_pedidos', pd.DataFrame())
        df_sab = st.session_state.get('df_sabana', pd.DataFrame())
        df_mez = st.session_state.get('df_mezclas', pd.DataFrame())
        df_cfg = st.session_state.get('df_config_base', pd.DataFrame())
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame())

        # Limpieza estricta de Finca para evitar clones
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
        mult_material = 1.112; tarifa_serv_tec_base = 1337.0; mult_avion = 1.112
        if not df_cfg.empty:
            match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
            if not match_cfg.empty:
                fila_c = match_cfg.iloc[0]
                mult_material = extraer_numero(fila_c.iloc[3])
                tarifa_serv_tec_base = extraer_numero(fila_c.iloc[4])
                mult_avion = extraer_numero(fila_c.iloc[6])

        # --- C. 🚀 CAZADOR DE DÍAS CICLO (FILTRO LÁSER) ---
        dias_ciclo_calc = 0
        
        if not df_apoyo.empty:
            col_finca = [c for c in df_apoyo.columns if 'FINCA' in str(c).upper()]
            col_fecha = [c for c in df_apoyo.columns if 'FECHA' in str(c).upper()]
            
            if col_finca and col_fecha:
                # FILTRO ESTRICTO: SACRAMENTO 1 != SACRAMENTO 10
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
                texto_pedido = match_ped.to_string().upper()
                for p_val in lista_pistas_validas:
                    if p_val in texto_pedido: pista_detectada = p_val; break
                
                for _, r_p in match_ped.iterrows():
                    if len(r_p) >= 7:
                        val_mat = str(r_p.iloc[5]).strip()
                        if "459" in val_mat:
                            ha_dosis_detectada = extraer_numero(r_p.iloc[6])
                            break

        ha_cobro_detectada = extraer_numero(datos_raw.get(8, 0))
        if ha_dosis_detectada == 0: ha_dosis_detectada = ha_cobro_detectada

        # --- 2. PANEL CONTROLES ---
        casilla_key = f"{finca_sel}_{vuelo_ref}_{fecha_operacion}"
        with st.container(border=True):
            st.markdown("#### ⚙️ Parámetros de Operación e Inteligencia de Ciclos")
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            r1c1.info(f"🧑‍🌾 Productor: **{tipo_productor}**")
            r1c2.warning(f"🛣️ Tope Finca: **{tipo_de_tope_finca}**")
            
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
        else:
            st.warning("🚨 No se encontró un pedido válido para la matriz de químicos.")
            costo_mezcla_total = 0.0

        # --- 4. TOPES (PRECIOS EXACTOS RESTAURADOS) ---
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

        if st.button("💾 DETONAR FACTURA Y GUARDAR", type="primary", use_container_width=True):
            st.balloons()
            st.success(f"Operación de {finca_sel} guardada con éxito.")
