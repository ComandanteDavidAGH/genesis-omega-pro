import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    """ Centraliza la autenticación con Google Cloud una sola vez en RAM """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=3600)
def obtener_radar_precios_cached(_extraer_numero):
    """ Función ultrarrápida vectorizada para leer y estructurar el radar de precios """
    gc = inicializar_cliente_gspread()
    if gc is None:
        return False, "🚨 Enlace satelital roto con Google Cloud.", None, 0, 0, 0, None
        
    try:
        sh = gc.open_by_url(URL_BOVEDA_MAESTRA)
        ws_conf = sh.worksheet("Configuración")
        
        data = ws_conf.get_all_values()
        df_conf = pd.DataFrame(data[1:], columns=data[0])
        
        radar = df_conf.iloc[:, [8, 9, 10]].copy()
        radar.columns = ['PRODUCTO', 'PRECIO_ACTUAL', 'PRECIO_SAP']
        
        # ⚡ Vectorización pura: Eliminar filas basura sin usar bucles for
        mask_validos = (
            radar['PRODUCTO'].notna() & 
            (radar['PRODUCTO'].str.strip() != "") & 
            (radar['PRODUCTO'].str.upper() != "PRODUCTO") &
            (radar['PRODUCTO'].str.upper() != "NAN") &
            (radar['PRODUCTO'].str.upper() != "NONE") &
            (radar['PRODUCTO'].str.strip() != "0") &
            (radar['PRODUCTO'].str.strip() != "0.0")
        )
        radar = radar[mask_validos].copy()
        
        # Aplicación vectorizada de conversión numérica
        radar['PRECIO_ACTUAL'] = radar['PRECIO_ACTUAL'].apply(_extraer_numero)
        radar['PRECIO_SAP'] = radar['PRECIO_SAP'].apply(_extraer_numero)
        radar['DIFERENCIA'] = (radar['PRECIO_SAP'] - radar['PRECIO_ACTUAL']).round(3)
        
        # Condicional vectorizado ultra-rápido (np.where)
        import numpy as np
        radar['ESTADO'] = np.where(radar['DIFERENCIA'].abs() < 0.001, "✅ OK", "❌ DESFASE")
        radar = radar.sort_values(by="ESTADO", ascending=False)
        
        total_insumos = len(radar)
        insumos_ok = len(radar[radar['ESTADO'] == "✅ OK"])
        insumos_fail = len(radar[radar['ESTADO'] == "❌ DESFASE"])
        
        return True, "Radar procesado", radar, total_insumos, insumos_ok, insumos_fail, data_full_export(data)
    except Exception as e:
        return False, f"🚨 Error al leer radar: {e}", None, 0, 0, 0, None

def data_full_export(data):
    """Función de apoyo auxiliar para no perder la matriz full"""
    return data

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE PRECIOS SAP
# =================================================================

def ejecutar(extraer_numero):
    st.markdown("""
    <style>
    .titulo-principal { 
        color: #0d1b2a; 
        border-bottom: 3px solid #d4af37; 
        padding-bottom: 5px; 
        font-family: 'Arial Black', sans-serif; 
    }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { 
        border: 3px solid #0d1b2a !important; 
        border-radius: 8px !important; 
        overflow: hidden !important; 
    }
    
    /* HUD de Control de Precios */
    .hud-precios {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .hud-precios-item { text-align: center; flex: 1; }
    .hud-precios-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-precios-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    .hud-precios-ok { color: #00ff66; font-family: 'Arial Black'; }
    .hud-precios-fail { color: #ff3333; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>🛠️ Mantenimiento Plantilla SAP</h1>", unsafe_allow_html=True)
     
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
                        except Exception:
                            f_sap_raw.seek(0)
                            df = pd.read_csv(f_sap_raw, sep=None, engine='python', encoding='latin1')
                     
                    # =========================================================
                    # 🛡️ MEJORA: ESCUDO ANTI-ARCHIVOS INVÁLIDOS
                    # =========================================================
                    if df.empty or len(df.columns) < 11:
                        st.error("🚨 ARCHIVO INVÁLIDO: La matriz cargada no tiene la estructura de SAP requerida (mínimo 11 columnas). Operación abortada.")
                        st.stop()
                        
                    df = df.dropna(subset=[df.columns[0]])
                    df = df[~df.iloc[:, 0].astype(str).str.contains(r'\*')]
                    df = df.sort_values(by=df.columns[10], ascending=True)
                     
                    df_final = df.iloc[:, 0:9].copy()
                    df_final['J'] = df.iloc[:, 10].values
                    unicos = sorted(df.iloc[:, 10].astype(str).unique().tolist())
                     
                    gc = inicializar_cliente_gspread()
                    if gc is None:
                        st.error("🚨 No se pudo establecer conexión con Google Cloud. Verifique sus credenciales.")
                        st.stop()
                        
                    boveda = gc.open_by_url(URL_BOVEDA_MAESTRA)
                    hoja_plantilla = boveda.worksheet("Plantilla")
                    hoja_plantilla.batch_clear(["A3:K5000"])
                    hoja_plantilla.update(range_name="A3", values=df_final.fillna("").values.tolist(), value_input_option='USER_ENTERED')
                    hoja_plantilla.update(range_name="K3", values=[[x] for x in unicos], value_input_option='USER_ENTERED')

                    st.success("✅ PASO A COMPLETADO: Datos frescos cargados en Plantilla de Drive de forma instantánea.")
                    st.session_state['paso_a_listo'] = True
                    st.cache_data.clear()
                    
                except Exception as e:
                    st.error(f"🚨 Error Crítico en Paso A al leer o procesar el archivo: {e}")

    st.markdown("---")
    st.markdown("### ⚡ PASO B: SINCRONIZADOR DE PRECIOS (ESTADO DEL ARSENAL)")

    def inyectar_precios_a_supabase(data_full):
        if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
            return False, "❌ ERROR: No hay conexión con la base de datos de Supabase."
        try:
            cliente_sb = st.session_state['supabase']
            dict_unicos = {}
            for fila in data_full[1:]:
                prod = fila[8] if len(fila) > 8 else ""
                val_k = fila[10] if len(fila) > 10 else ""
                if prod and str(prod).strip() and str(prod).upper() != "PRODUCTO":
                    prod_limpio = str(prod).strip().upper()
                    if prod_limpio not in ["0", "0.0", "NAN", "NONE", "NULL", ""]:
                        dict_unicos[prod_limpio] = str(val_k).strip()

            records_espejo = [
                {
                    "PRODUCTO": k, 
                    "COSTO": v,
                    "Columna2": "",
                    "valor a devolver": ""
                } for k, v in dict_unicos.items()
            ]

            if records_espejo:
                # 🛡️ MEJORA: UPSERT SEGURO EN LUGAR DE DELETE FANTASMA
                res = cliente_sb.table("PRECIOS_INSUMOS").upsert(records_espejo).execute()
                if res.data or res.data == []: 
                    return True, f"✅ Supabase actualizado (Upsert Seguro) con {len(records_espejo)} insumos."
                    
            return False, "⚠️ No se encontraron registros válidos para sincronizar."
        except Exception as e:
            return False, f"🚨 Error en Supabase durante Upsert: {e}"

    col_scan1, col_scan2 = st.columns([1, 1])
    
    with col_scan1:
        if st.button("🔍 ESCANEAR ESTADO ACTUAL", use_container_width=True):
            with st.spinner("Escaneando el estado de la bóveda de precios..."):
                ok, msg, radar, total_insumos, insumos_ok, insumos_fail, data_full = obtener_radar_precios_cached(extraer_numero)
                if not ok:
                    st.error(msg)
                else:
                    st.session_state['scan_ejecutado'] = True
                    st.session_state['radar_data'] = radar
                    st.session_state['total_insumos'] = total_insumos
                    st.session_state['insumos_ok'] = insumos_ok
                    st.session_state['insumos_fail'] = insumos_fail
                    st.session_state['data_full_cache'] = data_full
                    st.rerun()

    if st.session_state.get('scan_ejecutado'):
        radar = st.session_state.get('radar_data', pd.DataFrame())
        total_insumos = st.session_state.get('total_insumos', 0)
        insumos_ok = st.session_state.get('insumos_ok', 0)
        insumos_fail = st.session_state.get('insumos_fail', 0)
        data_full = st.session_state.get('data_full_cache', [])

        st.markdown(f"""
        <div class="hud-precios">
            <div class="hud-precios-item">
                <p class="hud-precios-title">Insumos Mapeados</p>
                <p class="hud-precios-value">🧪 {total_insumos}</p>
            </div>
            <div class="hud-precios-item">
                <p class="hud-precios-title">Nivel Estabilizado</p>
                <p class="hud-precios-value hud-precios-ok">🟢 {insumos_ok} OK</p>
            </div>
            <div class="hud-precios-item">
                <p class="hud-precios-title">Desfases Detectados</p>
                <p class="hud-precios-value {'hud-precios-fail' if insumos_fail > 0 else 'hud-precios-ok'}">
                    {'⚠️' if insumos_fail > 0 else '✅'} {insumos_fail} Desfases
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("#### 🛰️ Reporte Detallado de Situación:")

        c_btn1, c_btn2, c_btn3 = st.columns([1.5, 1.5, 1])
        
        with c_btn1:
            if st.button("🚀 NIVELAR Y SINCRONIZAR TODO (DRIVE + SUPABASE)", type="primary", use_container_width=True):
                with st.spinner("Nivelando Google Drive y Supabase Cloud..."):
                    try:
                        gc = inicializar_cliente_gspread()
                        if gc is None:
                            st.error("🚨 Enlace satelital roto con Google Drive.")
                            st.stop()
                            
                        sh = gc.open_by_url(URL_BOVEDA_MAESTRA)
                        ws_conf = sh.worksheet("Configuración")
                        
                        if not data_full:
                            data_full = ws_conf.get_all_values()
                        
                        valores_para_j = []
                        for fila in data_full[1:]:
                            valor_k = fila[10] if len(fila) > 10 else ""
                            valores_para_j.append([valor_k])
                         
                        if valores_para_j:
                            rango_destino = f"J2:J{len(valores_para_j) + 1}"
                            ws_conf.update(range_name=rango_destino, values=valores_para_j, value_input_option='USER_ENTERED')
                            st.toast("✅ Google Drive actualizado.", icon="📊")

                        ok_sb, msg_sb = inyectar_precios_a_supabase(data_full)
                        if ok_sb:
                            st.toast(msg_sb, icon="🌩️")
                        else:
                            st.warning(msg_sb)

                        st.cache_data.clear()
                        ok_re, msg_re, radar_re, tot, ok_i, fail_i, df_re = obtener_radar_precios_cached(extraer_numero)
                        if ok_re:
                            st.session_state['radar_data'] = radar_re
                            st.session_state['total_insumos'] = tot
                            st.session_state['insumos_ok'] = ok_i
                            st.session_state['insumos_fail'] = fail_i
                            st.session_state['data_full_cache'] = df_re
                            st.toast("⚡ ¡AMBAS NUBES SINCRONIZADAS AL 100%!", icon="✅")
                            st.balloons()
                            st.rerun()

                    except Exception as e:
                        st.error(f"🚨 Error durante la sincronización: {e}")

        with c_btn2:
            if st.button("🌩️ SINCRONIZAR ÚNICAMENTE SUPABASE CLOUD", use_container_width=True):
                with st.spinner("Conectando con la base relacional Supabase..."):
                    try:
                        if not data_full:
                            gc = inicializar_cliente_gspread()
                            if gc is None:
                                st.error("🚨 Enlace satelital roto con Google Drive.")
                                st.stop()
                            sh = gc.open_by_url(URL_BOVEDA_MAESTRA)
                            ws_conf = sh.worksheet("Configuración")
                            data_full = ws_conf.get_all_values()

                        ok_sb, msg_sb = inyectar_precios_a_supabase(data_full)
                        if ok_sb:
                            st.success(f"⚡ {msg_sb}")
                            st.balloons()
                        else:
                            st.error(msg_sb)

                    except Exception as e:
                        st.error(f"🚨 Error de conexión a Supabase: {e}")

        with c_btn3:
            if st.button("🔄 RE-ESCANEAR BÓVEDA", use_container_width=True):
                with st.spinner("Re-escaneando bóveda de precios..."):
                    st.cache_data.clear()
                    ok_re, msg_re, radar_re, tot, ok_i, fail_i, df_re = obtener_radar_precios_cached(extraer_numero)
                    if ok_re:
                        st.session_state['radar_data'] = radar_re
                        st.session_state['total_insumos'] = tot
                        st.session_state['insumos_ok'] = ok_i
                        st.session_state['insumos_fail'] = fail_i
                        st.session_state['data_full_cache'] = df_re
                    st.rerun()

        def color_estado(val):
            if val == "✅ OK": return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'
            if val == "❌ DESFASE": return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center;'
            return ''

        if not radar.empty:
            st.dataframe(
                radar.style.map(color_estado, subset=['ESTADO']), 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "PRECIO_ACTUAL": st.column_config.NumberColumn("PRECIO ACTUAL", format="%.3f"),
                    "PRECIO_SAP": st.column_config.NumberColumn("PRECIO SAP", format="%.3f"),
                    "DIFERENCIA": st.column_config.NumberColumn("DIFERENCIA", format="%.3f")
                }
            )
             
            if insumos_fail == 0:
                st.success("🟢 TODO EL SISTEMA ESTÁ EN NIVEL 'OK'. No se requieren ajustes operacionales.")
            else:
                st.warning("⚠️ SE DETECTARON DESFASES EN EL ARSENAL DE PRECIOS. Use los botones superiores para nivelar los tableros.")

if __name__ == "__main__":
    pass
