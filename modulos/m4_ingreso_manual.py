import streamlit as st
import pandas as pd
import gspread
import io
import re
from datetime import datetime, timedelta

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    """ Centraliza la autenticación con Google Cloud una sola vez en RAM """
    try:
        # 🌟 CORRECCIÓN MAESTRA: Cambiamos "gcp_credentials" por "gcp_service_account"
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

# --- 🧪 TRADUCTOR SEGURO DE NÚMEROS ---
def a_numero_limpio(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip().replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1:
            partes = v.rsplit('.', 1)
            v = partes[0].replace('.', '') + '.' + partes[1]
        return float(v) if v else 0.0
    except: return 0.0

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DEL MÓDULO DE ÓRDENES (OS)
# =================================================================

def ejecutar(extraer_numero, purificar_lote):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    
    /* HUD de Legalización */
    .hud-legalizador {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 12px; border-radius: 6px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.1); margin-bottom: 20px; text-align: center;
    }
    .hud-leg-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; }
    .hud-leg-value { font-size: 20px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Gestión y Legalización de Órdenes (OS)</h1>", unsafe_allow_html=True)
    
    # Cliente acelerado en RAM único para todo el archivo
    gc = inicializar_cliente_gspread()
    if gc is None:
        st.error("🚨 Enlace satelital roto con Google Cloud. Verifique sus credenciales.")
        return

    tab1, tab2 = st.tabs(["📝 1. Ingreso OS Manual (Desde Cero)", "🔄 2. Legalizar Vuelos Virtuales (Automático)"])

    # -----------------------------------------------------------------
    # PESTAÑA 1: INGRESO MANUAL ACELERADO (V3)
    # -----------------------------------------------------------------
    with tab1:
        st.subheader("Puesto de Control y Digitación Rápida")
        col_ref1, col_ref2 = st.columns([3, 1])
        with col_ref2:
            if st.button("🔄 RECARGAR BASES MANUALES", use_container_width=True, key="btn_recargar_m4"):
                st.session_state.pop('memoria_excel', None)
                st.rerun()

        try:
            boveda1 = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
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
            st.error(f"🚨 Error de enlace en Pestaña 1: {e}")
            st.stop()

        st.markdown("---")
        with st.expander("📝 1. DATOS DE LA ORDEN", expanded=True):
            c1, c2, c3 = st.columns(3)
            os_val = c1.text_input("Nº Orden (Ej: 318)", key="os_manual")
            fecha_dt = c2.date_input("📅 Fecha de Operación", format="DD/MM/YYYY", key="fecha_manual")
            piloto_val = c3.selectbox("👨‍✈️ Piloto", ["---"] + mem.get('lista_pilotos', []), key="piloto_manual")
            
            c4, c5, c6 = st.columns(3)
            hk_val = c4.selectbox("✈️ Matrícula (HK)", ["---"] + mem.get('lista_hks', []), key="hk_manual")
            horo_val = c5.text_input("⏱️ Horómetro TOTAL (Ej: 1.5)", value="0", key="horo_manual")
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
                        
                        mod_av, pist_av = "", ""
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
                            
                            bloq, sect, hab, t_prod = "", "", 0, ""
                            m_f = df_t2_m4[df_t2_m4.iloc[:, 0].str.upper().str.strip() == n_finca]
                            if not m_f.empty:
                                sect, hab, bloq, t_prod = m_f.iloc[0, 1], extraer_numero(m_f.iloc[0, 2]), m_f.iloc[0, 3], m_f.iloc[0, 5]
                            
                            coctel_final = str(f.get('coctel', '')).strip()
                            if not coctel_final or coctel_final in ["None", ""]:
                                mask = (df_apoyo_m4.iloc[:, 1].str.upper().str.strip() == n_finca) & (df_apoyo_m4.iloc[:, 5].str.strip() == f_str)
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
                            
                            # Fórmulas Inteligentes de Excel
                            row[24] = '=INDIRECT("Y"&(ROW()-1))'  
                            row[25] = '=INDIRECT("Z"&(ROW()-1))'  
                            row[26] = '=IFERROR(INDIRECT("S"&ROW())/INDIRECT("F"&ROW()), 0)' 
                            row[27] = '=IF(INDIRECT("AA"&ROW())>INDIRECT("Z"&ROW()), "SUPERIOR", "INFERIOR")' 
                            row[30] = '=INDIRECT("AE"&(ROW()-1))' 
                            
                            filas_finales.append(row)
                        
                        if filas_finales:
                            hoja_maestra1.append_rows(filas_finales, value_input_option='USER_ENTERED')
                            st.balloons()
                            st.success(f"🎯 ¡OPERACIÓN EXITOSA! OS {os_val} inyectada con Cóctel y Fórmulas Automáticas.")
                            st.session_state.pop('memoria_excel', None) 
                            st.rerun()
                                
                except Exception as e: st.error(f"Error en inyección: {e}")

    # -----------------------------------------------------------------
    # PESTAÑA 2: ESCÁNER DE LEGALIZACIÓN MULTI-OS (MEMORIA RAM ACELERADA)
    # -----------------------------------------------------------------
    with tab2:
        st.markdown("### 🔄 Escáner de Legalización Multi-OS")
        
        sh2 = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        ws_t1_2 = sh2.worksheet("TABLA 1")
        ws_apoyo_2 = sh2.worksheet("TABLA DE APOYO2023")
        
        # ⚡ Cacheamos los datos de red de las tablas para mejorar rendimiento
        if 'legalizacion_cache' not in st.session_state:
            with st.spinner("Escaneando bases generales para legalizar misiones virtuales..."):
                cache_leg = {}
                cache_leg['datos_t1'] = ws_t1_2.get_all_values()
                cache_leg['datos_apoyo'] = ws_apoyo_2.get_all_values()
                st.session_state['legalizacion_cache'] = cache_leg
                
        if st.button("🔄 RECARGAR VUELOS VIRTUALES", use_container_width=True):
            st.session_state.pop('legalizacion_cache', None)
            st.rerun()

        datos_t1 = st.session_state['legalizacion_cache']['datos_t1']
        datos_apoyo = st.session_state['legalizacion_cache']['datos_apoyo']
        
        pendientes = []
        for idx, row in enumerate(datos_t1[5:]):
            if len(row) > 19:
                os_val_check = str(row[0]).upper()
                equipo = str(row[17]).upper() 
                if os_val_check.startswith("VIRT-") and ("AVION" in equipo or equipo == ""):
                    pendientes.append({
                        "fila_real": idx + 6, "os_virt": os_val_check, "finca": row[2],
                        "ha": extraer_numero(row[5]), "costo_ha": extraer_numero(row[19]), 
                        "total": extraer_numero(row[18]), "modelo": equipo
                    })

        if not pendientes:
            st.success("✅ No hay misiones de Avión pendientes por legalizar. ¡Cielo despejado!")
        else:
            df_pend = pd.DataFrame(pendientes)
            opciones_virt = df_pend.apply(lambda x: f"Fila {x['fila_real']} | {x['finca']} | {x['ha']} Ha | {x['os_virt']}", axis=1).tolist()
            
            st.markdown(f"""
            <div class="hud-legalizador">
                <div class="hud-precios-item">
                    <p class="hud-leg-title">Misiones Virtuales en Espera</p>
                    <p class="hud-leg-value">🛰️ {len(pendientes)} Órdenes por Legalizar</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            seleccion = st.selectbox("🎯 Seleccione Vuelo Virtual para Legalizar:", opciones_virt)
            vuelo_sel = df_pend.iloc[opciones_virt.index(seleccion)]
            
            st.markdown("---")
            st.subheader(f"🛠️ Desglose de OS para: {vuelo_sel['finca']}")
            
            lista_todas_fincas = sorted(list(set([r[1] for r in datos_apoyo[3:] if len(r) > 1 and r[1]])))

            if 'legalizador_rows' not in st.session_state:
                st.session_state.legalizador_rows = [{"OS_Real": "", "Finca": vuelo_sel['finca'], "Hect AREAS": vuelo_sel['ha'], "Costo_Ha": vuelo_sel['costo_ha']}]

            col_btn, _ = st.columns([1, 4])
            if col_btn.button("➕ Añadir Finca/OS al Combo"):
                st.session_state.legalizador_rows.append({"OS_Real": "", "Finca": "", "Hect AREAS": 0.0, "Costo_Ha": 0.0})
                st.rerun()

            rows_finales = []
            total_ha_asignadas = 0.0

            for i, row in enumerate(st.session_state.legalizador_rows):
                with st.container(border=True):
                    c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                    os_r = c1.text_input(f"OS Real #{i+1}", value=row["OS_Real"], key=f"os_r_{i}")
                    finca_r = c2.selectbox(f"Finca #{i+1}", [""] + lista_todas_fincas, index=lista_todas_fincas.index(row["Finca"])+1 if row["Finca"] in lista_todas_fincas else 0, key=f"f_r_{i}")
                    
                    costo_sugerido = row["Costo_Ha"]
                    if finca_r != row["Finca"] and finca_r != "":
                        for r_ap in reversed(datos_apoyo):
                            if len(r_ap) > 3 and r_ap[1] == finca_r:
                                costo_sugerido = extraer_numero(r_ap[3])
                                break
                    
                    ha_r = c3.number_input(f"Ha #{i+1}", value=float(row.get("Hect AREAS", 0.0)) if "Hect AREAS" in row else float(row.get("Hectáreas", 0.0)), key=f"h_r_{i}")
                    costo_r = c4.number_input(f"$/Ha #{i+1}", value=float(costo_sugerido), key=f"c_r_{i}")
                    
                    rows_finales.append({"OS": os_r, "Finca": finca_r, "Ha": ha_r, "Costo": costo_r})
                    if finca_r == vuelo_sel['finca']: total_ha_cobro_escuadron = ha_r 
                    total_ha_asignadas += ha_r

            st.markdown("---")
            diferencia = round(vuelo_sel['ha'] - total_ha_asignadas, 2)
            
            c_m1, c_m2 = st.columns(2)
            c_m1.metric("🚜 Ha Objetivo (Finca Original)", f"{vuelo_sel['ha']} Ha")
            c_m2.metric("⚖️ Diferencia Pendiente", f"{diferencia} Ha", delta=-diferencia, delta_color="inverse")

            if st.button("🚀 DETONAR LEGALIZACIÓN EN TABLA 1", type="primary", use_container_width=True):
                if abs(diferencia) > 0.05:
                    st.error(f"❌ Error de cuadre: Aún faltan {diferencia} Ha por asignar en el manifiesto real de vuelo.")
                else:
                    try:
                        with st.spinner("Legalizando y respetando Fórmulas MAP de Excel..."):
                            r_idx = int(vuelo_sel['fila_real'])
                            Nuevas_Filas = []
                            
                            for r_f in rows_finales:
                                fila_orig = datos_t1[r_idx - 1]
                                nueva = list(fila_orig) 
                                
                                nueva[0] = str(r_f["OS"])       
                                nueva[2] = str(r_f["Finca"])    
                                nueva[5] = float(r_f["Ha"])       
                                nueva[19] = float(r_f["Costo"])   
                                nueva[18] = float(round(r_f["Ha"] * r_f["Costo"], 0)) 
                                nueva[21] = nueva[18]      
                                
                                indices_vacios = [24, 25, 26, 27, 30]
                                for idx_v in indices_vacios:
                                    if idx_v < len(nueva): nueva[idx_v] = ""
                                
                                Nuevas_Filas.append(nueva)

                            ws_t1_2.delete_rows(r_idx)
                            ws_t1_2.insert_rows(Nuevas_Filas, r_idx, value_input_option='USER_ENTERED')
                            
                            st.balloons()
                            st.success(f"🎯 LEGALIZACIÓN PERFECTA. El registro virtual ha sido eliminado y reemplazado por misiones reales.")
                            
                            st.session_state.pop('legalizador_rows', None)
                            st.session_state.pop('legalizacion_cache', None)
                            st.rerun()
                    except Exception as e:
                        st.error(f"🚨 Falla en el sistema de inserción de filas: {e}")

if __name__ == "__main__":
    ejecutar(None, None)
