import streamlit as st
import pandas as pd
import gspread
import requests
import io
import re
import math
from datetime import datetime

def ejecutar(extraer_numero, fmt_sap, procesar_fecha_pesada):
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
                    col_mat = [c for c in df_p.columns if 'MATERIAL' in str(c).upper() or 'ITEM' in str(c).upper()][0]
                    
                    finca_sap = str(match_sap.iloc[0][col_finca]).strip().upper()
                    
                    # 🎯 REGLA DE ORO 459: Francotirador directo a la columna Material
                    ha_correcta = 0.0
                    for _, fila_ped in match_sap.iterrows():
                        valor_material = str(fila_ped[col_mat]).strip()
                        if valor_material == "459" or valor_material.split(".")[0] == "459": 
                            ha_correcta = extraer_numero(fila_ped[col_ha])
                            break
                    
                    if ha_correcta > 0:
                        st.session_state['ha_radar_sap'] = ha_correcta
                    else:
                        st.session_state['ha_radar_sap'] = extraer_numero(match_sap.iloc[0][col_ha])
                    
                    st.success(f"✅ **SAP CONFIRMADO:** {finca_sap} | {st.session_state['ha_radar_sap']} Ha")
                except:
                    pass

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
        datos_raw = datos_vuelo.get('DATOS_FILA', {})
        
        num_pedido = "S/N"
        if pedido_sap and len(str(pedido_sap)) >= 7:
            num_pedido = str(pedido_sap).strip()
        elif datos_vuelo.get('PEDIDO_SAP') and str(datos_vuelo.get('PEDIDO_SAP')).strip() != "":
            num_pedido = str(datos_vuelo.get('PEDIDO_SAP')).strip()
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
            
            mision_solo_dron = c_sup2.toggle("🤖 MISIÓN 100% DRON", value=False, key=f"dron_toggle_{casilla_key}")
            
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            r1c1.number_input("📅 Ciclo (SISTEMA)", value=int(dias_ciclo_calc), disabled=True, key=f"ds_{casilla_key}")
            d_ciclo_factura = r1c2.number_input("⏳ Ciclo (COBRO)", value=int(dias_ciclo_calc), step=1, key=f"df_{casilla_key}")
            
            ha_sugerida = float(st.session_state.get('ha_radar_sap', 0.0))
            if ha_sugerida == 0.0: ha_sugerida = float(ha_dosis_detectada)
                
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
                recargo_lista = r2c2.selectbox("Cargo Terrestre:", opciones_rec, index=idx_recargo, key=f"rl_{casilla_key}")
                if recargo_lista == "Otro Valor Manual...":
                    recargo_final = r2c3.number_input("✍️ Digite Recargo ($)", value=0, step=1000, key=f"rm_{casilla_key}")
                else:
                    recargo_final = float(recargo_lista.split(" ")[0])

        dict_topes_pista = {"TOPE MAX GENERAL": {"PLUC": 63326, "PORI": 62718, "TEHO": 63325, "PDIV": 63325, "LUCI": 63325}, "TOPE SUR": {"PLUC": 71517, "PORI": 70829, "TEHO": 71517, "PDIV": 71517, "LUCI": 71517}, "TOPE PARCELA INTER < 20HA": {"PLUC": 98335, "PORI": 105723, "TEHO": 98335, "PDIV": 105723, "LUCI": 98335}}
        val_tope = dict_topes_pista.get(tipo_de_tope_finca, {}).get(pista_sel, 999999)
        
        with st.container(border=True):
            st.markdown("#### ✈️ Hangar de Despliegue")
            costo_total_vuegos = 0.0
            costo_neto_vuelo_total = 0.0  
            total_ha_cobro_escuadron = 0.0
            horometro_final_avion = 0.0 

            if mision_solo_dron:
                st.success("🚁 Modo Dron Activo: Costos calculados sin recargos terrestres ni topes de pista.")
                try:
                    if "gcp_credentials" in st.secrets:
                        gc_vd = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc_vd = gspread.service_account(filename='credenciales.json')
                    boveda_vd = gc_vd.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    datos_vd = boveda_vd.worksheet("Validación Dosis").get_all_values()
                    df_flota = pd.DataFrame(datos_vd[2:], columns=datos_vd[1]) 
                    df_dr = df_flota[df_flota['Tarifa'].notna() & (df_flota['Tarifa'].astype(str).str.strip() != '')]
                    nombres_dr = df_dr['Tarifa'].astype(str).str.replace('TARIFA ', '', case=False).str.strip()
                    nombres_dr = nombres_dr.apply(lambda x: f"DRONE {x}" if "DRONE" not in x.upper() else x)
                    precios_dr = pd.to_numeric(df_dr['Valor ha/Dr'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)
                    dict_drones = dict(zip(nombres_dr, precios_dr))
                except Exception as e:
                    dict_drones = {"DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280}

                df_drones_def = pd.DataFrame([{"Drone": "DRONE DATAROT", "Hectáreas": float(ha_cobro_detectada)}])
                escuadron_drones = st.data_editor(df_drones_def, key=f"drones_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=list(dict_drones.keys()), required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f", required=True)}, use_container_width=True, hide_index=True)
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                    costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                    costo_total_vuegos += (tarifa_dron_neta * ha_dr) * mult_avion_final 

            else:
                c_av, c_dr = st.columns(2)
                try:
                    if "gcp_credentials" in st.secrets:
                        gc_vd = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc_vd = gspread.service_account(filename='credenciales.json')
                    boveda_vd = gc_vd.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    datos_vd = boveda_vd.worksheet("Validación Dosis").get_all_values()
                    df_flota = pd.DataFrame(datos_vd[2:], columns=datos_vd[1]) 
                    df_av = df_flota[df_flota['TIPO'].notna() & (df_flota['TIPO'].astype(str).str.strip() != '')]
                    dict_aviones = dict(zip(df_av['TIPO'].astype(str).str.strip(), pd.to_numeric(df_av['HORA'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)))
                    df_dr = df_flota[df_flota['Tarifa'].notna() & (df_flota['Tarifa'].astype(str).str.strip() != '')]
                    nombres_dr = df_dr['Tarifa'].astype(str).str.replace('TARIFA ', '', case=False).str.strip()
                    nombres_dr = nombres_dr.apply(lambda x: f"DRONE {x}" if "DRONE" not in x.upper() else x)
                    precios_dr = pd.to_numeric(df_dr['Valor ha/Dr'].astype(str).str.replace('.', '', regex=False), errors='coerce').fillna(0)
                    dict_drones = dict(zip(nombres_dr, precios_dr))
                except Exception as e:
                    dict_aviones = {"THRUS SR2": 4606562, "PIPER PA 36-375": 3985831, "CESSNA O PIPER PA 25": 3036525, "AIR TRACTOR": 4665109, "CESSNA ASA": 3666600, "CESSNA FUMIGARAY": 3065952}
                    dict_drones = {"DRONE DATAROT": 84428, "DRONE NORTE": 75518, "DRONE AVIL": 71280, "DRONE GENESYS": 71280}

                with c_av: 
                    st.markdown("##### 🛩️ Base Aviones")
                    df_aviones_def = pd.DataFrame([{"Avión": "CESSNA ASA", "Hectáreas": float(ha_cobro_detectada), "Horómetro": 1.00}])
                    opciones_av = list(dict_aviones.keys()) if 'dict_aviones' in locals() and dict_aviones else ["THRUS SR2", "PIPER PA 36-375"]
                    escuadron_aviones = st.data_editor(df_aviones_def, key=f"aviones_{casilla_key}", num_rows="dynamic", column_config={"Avión": st.column_config.SelectboxColumn("Modelo", options=opciones_av, required=True), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f"), "Horómetro": st.column_config.NumberColumn("Horómetro", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)
                    
                with c_dr:
                    st.markdown("##### 🚁 Base Drones (Apoyo)")
                    df_drones_def = pd.DataFrame([{"Drone": None, "Hectáreas": 0.0}])
                    opciones_dr = list(dict_drones.keys()) if 'dict_drones' in locals() and dict_drones else ["DRONE DATAROT", "DRON GENESYS"]
                    escuadron_drones = st.data_editor(df_drones_def, key=f"drones_mix_{casilla_key}", num_rows="dynamic", column_config={"Drone": st.column_config.SelectboxColumn("Modelo Dron", options=opciones_dr), "Hectáreas": st.column_config.NumberColumn("Hectáreas", min_value=0.00, format="%.2f")}, use_container_width=True, hide_index=True)                
                
                # 🛡️ CÁLCULOS PROTEGIDOS E INTELIGENCIA NETA
                for index, row in escuadron_aviones.iterrows():
                    av_sel = row["Avión"]
                    try: ha_av = float(row.get("Hectáreas", 0)) if str(row.get("Hectáreas", 0)) not in ["None", "", "nan"] else 0.0
                    except: ha_av = 0.0
                        
                    try: horo = float(row.get("Horómetro", 0)) if str(row.get("Horómetro", 0)) not in ["None", "", "nan"] else 0.0
                    except: horo = 0.0
                    
                    if pd.isna(av_sel) or ha_av <= 0: continue
                    total_ha_cobro_escuadron += ha_av
                    horometro_final_avion += horo  
                    
                    tarifa_base_ha = (dict_aviones.get(av_sel, 0) * horo) / ha_av if ha_av > 0 else 0
                    tarifa_base_tope = tarifa_base_ha if pista_sel == "PDIV" else min(tarifa_base_ha, val_tope)
                    
                    costo_neto_vuelo_total += (tarifa_base_tope * ha_av) 
                    tarifa_aplicada = tarifa_base_tope + recargo_final
                    costo_total_vuegos += (tarifa_aplicada * ha_av) * mult_avion_final 
                    
                for _, row in escuadron_drones.iterrows():
                    dr_sel, ha_dr = row["Drone"], float(row.get("Hectáreas", 0))
                    if pd.isna(dr_sel) or ha_dr <= 0: continue
                    total_ha_cobro_escuadron += ha_dr
                    tarifa_dron_neta = dict_drones.get(dr_sel, 0)
                    costo_neto_vuelo_total += (tarifa_dron_neta * ha_dr)  
                    costo_total_vuegos += (tarifa_dron_neta * ha_dr) * mult_avion_final
            
        st.markdown("#### 🧪 Matriz de Validación e Inteligencia de Mezcla")
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
                sap_dict_pista[nombre_limpio] = sap_dict_pista.get(nombre_limpio, 0.0) + dosis_pista
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
                    if iter_id == coctel_piloto_base:
                        puntaje += 10000

                    for p_receta, d_esperada in receta.items():
                        match_receta = False
                        dose_matched = False
                        for k_sap, d_sap in sap_dict_pista.items():
                            if p_receta == k_sap or (len(k_sap)>=4 and p_receta in k_sap) or (len(p_receta)>=4 and k_sap in p_receta):
                                match_receta = True
                                if abs(d_sap - d_esperada) <= 0.5: dose_matched = True 
                                break
                        if match_receta: puntaje += 50 if dose_matched else 10
                        else: es_valido = False; break

                if es_valido and puntaje > max_p:
                    max_p = puntaje
                    coctel_base = iter_id
                    dosis_oficiales_coctel = receta.copy()
                    
            sigla_fertilizante = ""
            for k_sap in sap_dict_pista.keys():
                for f_name, f_sigla in dict_fertilizantes.items():
                    if f_name == k_sap or (len(k_sap)>=4 and f_name in k_sap) or (len(f_name)>=4 and k_sap in f_name):
                        sigla_fertilizante = f" {f_sigla}"
                        break
                if sigla_fertilizante: break

            coctel_ganador = coctel_base + sigla_fertilizante if coctel_base != "SIN COINCIDENCIA" else "SIN COINCIDENCIA"
            st.success(f"🤖 **MOTOR IA MAESTRO:** Cóctel Oficial: **{coctel_ganador}**")

            st.error(f"🔍 **RAYOS X DE COLUMNAS:** Pista Elegida: '{pista_sel}' | Lote: Columna {idx_lote} | Almacén: Columna {idx_almacen} | Saldo: Columna {idx_saldo}")
            
            matriz_datos = []
            for item_data in datos_extraidos_sap:
                cod_item = str(item_data['cod']).strip().upper().lstrip('0')
                nombre_p = item_data['nombre']
                nombre_limpio = item_data['nombre_limpio']
                cant_total_pedido = item_data['cant_total']

                costo_unit = 0.0; lote_sap = "SIN LOTE EN PISTA"; saldo_sap = 0.0

                if not df_sab.empty:
                    col_0_limpia = df_sab.iloc[:, 0].apply(lambda x: str(x).split('.')[0].strip().upper().lstrip('0') if str(x).lower() not in ['nan', 'none', '<na>', ''] else "")
                    match_sabana_global = df_sab[col_0_limpia == cod_item]
                    
                    if match_sabana_global.empty and nombre_limpio != "" and "ITEM" not in nombre_limpio:
                        match_sabana_global = df_sab[df_sab.astype(str).apply(lambda x: x.str.contains(nombre_limpio, case=False, na=False)).any(axis=1)]

                    if not match_sabana_global.empty:
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

                        if idx_almacen != -1:
                            col_almacen = match_sabana_global.iloc[:, idx_almacen].astype(str).str.strip().str.upper()
                            match_pista = match_sabana_global[col_almacen.str.contains(str(pista_sel).strip().upper(), na=False)]
                        else:
                            match_pista = match_sabana_global[match_sabana_global.astype(str).apply(lambda x: x.str.strip().str.upper().str.contains(str(pista_sel).strip().upper(), na=False)).any(axis=1)]

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

            df_matriz["Total_Fila"] = (df_matriz["I: Sugerido SAP (Total)"] * df_matriz["E: Costo Unit (+Margen)"])
            costo_mezcla_total = df_matriz["Total_Fila"].apply(lambda x: math.floor(x + 0.5)).sum()
            
            df_matriz = df_matriz.drop(columns=["B_Val", "C_Val", "Total_Fila"])

            def colorear_matriz(row):
                global_sap = df_matriz[df_matriz["A: Producto"] == row["A: Producto"]]["I: Sugerido SAP (Total)"].sum()
                ideal_sistema = row["D: Dosis Total (Sistema)"]
                diferencia = abs(global_sap - ideal_sistema)
                
                if diferencia <= 0.5:
                    color = 'background-color: #d4edda; color: #155724;' 
                elif diferencia <= 5.0:
                    color = 'background-color: #fff3cd; color: #856404;' 
                elif diferencia <= 20.0:
                    color = 'background-color: #f8d7da; color: #721c24;' 
                else:
                    color = 'background-color: #8b0000; color: white; font-weight: bold;' 
                    
                return [color] * len(row)

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
        
        def sap_round(n):
            return math.floor(n + 0.5)

        unitario_st_bruto = d_ciclo_factura * tarifa_serv_tec_base
        unitario_vuelo_bruto = costo_total_vuegos / total_ha_cobro_escuadron if total_ha_cobro_escuadron > 0 else 0
        
        unitario_st = sap_round(unitario_st_bruto)
        unitario_vuelo = sap_round(unitario_vuelo_bruto)
        
        subtotal_st_finca = sap_round(unitario_st * ha_dosis_final)
        subtotal_vuelo_finca = sap_round(unitario_vuelo * ha_dosis_final)
        
        gran_total = costo_mezcla_total + subtotal_vuelo_finca + subtotal_st_finca
        costo_por_ha = sap_round(gran_total / ha_dosis_final) if ha_dosis_final > 0 else 0

        precio_columna_referencia = 0
        if not mision_solo_dron:
            try:
                if not escuadron_aviones.empty:
                    avion_principal = escuadron_aviones.iloc[0]['Avión']
                    precio_columna_referencia = dict_aviones.get(avion_principal, 0)
            except:
                precio_columna_referencia = 0

        precio_dron_referencia = 0
        try:
            df_busqueda_dron = escuadron_drones if mision_solo_dron else escuadron_drones
            if not df_busqueda_dron.empty:
                dron_principal = df_busqueda_dron.iloc[0]['Drone']
                if pd.notna(dron_principal) and str(dron_principal).strip() not in ["None", ""]:
                    precio_dron_referencia = dict_drones.get(dron_principal, 0)
        except:
            precio_dron_referencia = 0

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
        with m2: 
            ha_avion_real = 0
            try:
                if not mision_solo_dron and not escuadron_aviones.empty:
                    ha_avion_real = float(escuadron_aviones['Hectáreas'].sum())
            except:
                ha_avion_real = 0

            es_dron_dominante = mision_solo_dron or (ha_avion_real == 0 and precio_dron_referencia > 0)
            st.markdown(mini_metric("🛣️", "Pista", tipo_de_tope_finca if not es_dron_dominante else "DRON"), unsafe_allow_html=True)
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True) 
            
            if es_dron_dominante:
                texto_valor_tope = f"$ {fmt_sap(precio_dron_referencia)}"
            elif val_tope == 999999 or val_tope == 0:
                texto_valor_tope = "Sin Tope"
            else:
                texto_valor_tope = f"$ {fmt_sap(val_tope)}"
                
            st.markdown(mini_metric("🚧", "Valor Tope", texto_valor_tope), unsafe_allow_html=True)
        with m3: st.markdown(mini_metric("👨‍🔬", "Tarifa ST", f"$ {fmt_sap(tarifa_serv_tec_base)}"), unsafe_allow_html=True)
        with m4: st.markdown(mini_metric("✈️", "Mult.", f"x {mult_avion_final}"), unsafe_allow_html=True)
        
        with m5: 
            st.markdown(mini_metric("⏱️", "Precio Hora", f"$ {fmt_sap(precio_columna_referencia)}"), unsafe_allow_html=True)
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

                    ha_f = float(ha_dosis_final)
                    
                    if mision_solo_dron:
                        h_total_v = ha_f / 10
                    else:
                        h_total_v = (ha_f / total_ha_cobro_escuadron) * horometro_final_avion if total_ha_cobro_escuadron > 0 else 0.0
                        
                    vol_total_gln = ha_f * 6
                    rend_min = h_total_v * 60
                    
                    piloto_f = "OPERADOR DRONE" if mision_solo_dron else "PILOTO AVIÓN"
                    
                    if mision_solo_dron:
                        if "DATAROT" in tipo_mision.upper(): hk_f = "DR51"
                        elif "GENESYS" in tipo_mision.upper(): hk_f = "DR52"
                        elif "AVIL" in tipo_mision.upper(): hk_f = "DR53"
                        else: hk_f = "DRONE_GEN"
                    else:
                        hk_f = hk_sel if 'hk_sel' in locals() else "AVION_REG"

                    modelo_f = "DRONE" if mision_solo_dron else "AVION"

                    try:
                        fila_excel = len(st.session_state['df_azul_actual']) + 2 
                    except:
                        fila_excel = 6 
                    
                    tarifa_vuelo_neta_ha = float(costo_neto_vuelo_total / total_ha_cobro_escuadron) if total_ha_cobro_escuadron > 0 else 0.0
                    valor_dominical = float(recargo_final)
                    
                    total_pago_avion_neto = (tarifa_vuelo_neta_ha + valor_dominical) * ha_f
                    
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
                    row_azul[10] = round(h_total_v, 2)
                    row_azul[11] = 6
                    row_azul[12] = round(vol_total_gln, 2) 
                    row_azul[13] = round(h_total_v, 2) 
                    row_azul[14] = round(rend_min, 2) 
                    row_azul[15] = piloto_f
                    row_azul[16] = hk_f
                    row_azul[17] = tipo_mision
                    
                    row_azul[18] = float(gran_total) 
                    row_azul[19] = round(tarifa_vuelo_neta_ha, 2) 
                    row_azul[20] = round(valor_dominical, 2) 
                    row_azul[21] = float(gran_total) 
                    row_azul[23] = pista_manual 
                    row_azul[28] = float(gran_total) 
                    row_azul[29] = round(total_pago_avion_neto, 2) 
                    row_azul[32] = tipo_productor 
                    row_azul[33] = "GÉNESIS_V2_PRO" 
                    
                    fila_apoyo = [""] * 15
                    fila_apoyo[0] = "=IFERROR(ROW()-3, 0)" 
                    fila_apoyo[1] = finca_limpia
                    fila_apoyo[2] = ha_f
                    fila_apoyo[3] = float(costo_por_ha)
                    fila_apoyo[5] = fecha_str
                    fila_apoyo[8] = coctel_ganador
                    fila_apoyo[10] = pista_manual
                    fila_apoyo[13] = tipo_mision
                    
                    col_azul = hoja_maestra.col_values(1)
                    fila_destino_azul = 1
                    for i in range(len(col_azul)-1, -1, -1):
                        if str(col_azul[i]).strip() != "":
                            fila_destino_azul = i + 2
                            break
                    
                    if fila_destino_azul > hoja_maestra.row_count:
                        hoja_maestra.add_rows(10)

                    col_apoyo = hoja_apoyo.col_values(2)
                    fila_destino_apoyo = 1
                    for i in range(len(col_apoyo)-1, -1, -1):
                        if str(col_apoyo[i]).strip() != "":
                            fila_destino_apoyo = i + 2
                            break
                            
                    if fila_destino_apoyo > hoja_apoyo.row_count:
                        hoja_apoyo.add_rows(10)

                    fila_apoyo[0] = fila_destino_apoyo - 3

                    hoja_maestra.update(range_name=f"A{fila_destino_azul}", values=[row_azul], value_input_option='USER_ENTERED')
                    hoja_apoyo.update(range_name=f"A{fila_destino_apoyo}", values=[fila_apoyo], value_input_option='USER_ENTERED')
                    
                    try:
                        datos_memoria = hoja_memoria.get_all_values()
                        set_existentes = set()
                        if len(datos_memoria) > 1:
                            for r in datos_memoria[1:]:
                                if len(r) >= 10:
                                    llave = f"{str(r[0]).strip()}|{str(r[9]).strip().upper()}|{str(r[3]).strip().upper()}"
                                    set_existentes.add(llave)
                        
                        bodega_f = "BODEGA PRINCIPAL DRON" if mision_solo_dron else "BODEGA PRINCIPAL AVIÓN"
                        
                        filas_memoria = []
                        contador_nuevos = 0
                        
                        for idx, row in edited_df.iterrows():
                            nombre_prod = str(row.get("A: Producto", "")).strip().upper()
                            if "⚠️" not in nombre_prod and nombre_prod != "" and nombre_prod != "NAN":
                                llave_actual = f"{fecha_str}|{finca_limpia}|{nombre_prod}"
                                
                                if llave_actual not in set_existentes:
                                    dosis_prod = float(row.get("D: Dosis Total (Sistema)", 0))
                                    lote_prod = str(row.get("G: Lotes (SAP)", "S/N"))
                                    
                                    fila_m = [""] * 10
                                    fila_m[0] = fecha_str 
                                    fila_m[1] = coctel_ganador 
                                    fila_m[2] = str(pista_manual).split("-")[0].strip()[:4] 
                                    fila_m[3] = nombre_prod 
                                    fila_m[4] = lote_prod 
                                    fila_m[5] = float(dosis_prod) 
                                    fila_m[6] = bodega_f 
                                    fila_m[7] = "" 
                                    fila_m[8] = "X" 
                                    fila_m[9] = finca_limpia 
                                    
                                    filas_memoria.append(fila_m)
                                    contador_nuevos += 1
                                    
                        if filas_memoria:
                            hoja_memoria.append_rows(filas_memoria, value_input_option='USER_ENTERED')
                            st.toast(f"💾 Memoria Samurai: {contador_nuevos} productos nuevos guardados.")
                        else:
                            st.toast("⚠️ Memoria Samurai: Los productos ya existían. No se duplicaron.")
                            
                    except Exception as e_mem:
                        st.warning(f"⚠️ Nota de sistema: Error al guardar en MEMORIA: {e_mem}")

                    st.balloons()
                    st.success(f"✅ IMPACTO TOTAL CONFIRMADO. Referencia {os_virtual} inyectada exactamente en la fila {fila_destino_azul}.")
                    
                    if 'memoria_excel' in st.session_state:
                        del st.session_state['memoria_excel']

                except Exception as e_save:
                    st.error(f"🚨 Falla en el Gatillo de Guardado: {e_save}")
