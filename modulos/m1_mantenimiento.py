import streamlit as st
import pandas as pd
import gspread
import io
import re
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    """ Centraliza la autenticación con Google Cloud una sola vez en RAM """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # 🌟 UNIFICACIÓN MAESTRA: Usamos el secreto gcp_service_account con scopes explícitos
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE PRECIOS SAP
# =================================================================

def ejecutar(extraer_numero):
    # Inyección de la línea estética VIP Corporativa
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
 
    st.markdown("<h1 class='titulo-principal'>Inteligencia de Precios SAP</h1>", unsafe_allow_html=True)
     
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
                        except:
                            f_sap_raw.seek(0)
                            df = pd.read_csv(f_sap_raw, sep=None, engine='python', encoding='latin1')
                     
                    df = df.dropna(subset=[df.columns[0]])
                    df = df[~df.iloc[:, 0].astype(str).str.contains('\*')]
                    if len(df.columns) >= 11:
                        df = df.sort_values(by=df.columns[10], ascending=True)
                     
                    df_final = df.iloc[:, 0:9].copy()
                    df_final['J'] = df.iloc[:, 10].values
                    unicos = sorted(df.iloc[:, 10].astype(str).unique().tolist())
                     
                    # Autenticación acelerada en RAM interconectada
                    gc = inicializar_cliente_gspread()
                    if gc is None:
                        st.error("🚨 No se pudo establecer conexión con Google Cloud. Verifique sus credenciales.")
                        st.stop()
                        
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    hoja_plantilla = boveda.worksheet("Plantilla")
                    hoja_plantilla.batch_clear(["A3:K5000"])
                    hoja_plantilla.update(range_name="A3", values=df_final.fillna("").values.tolist(), value_input_option='USER_ENTERED')
                    hoja_plantilla.update(range_name="K3", values=[[x] for x in unicos], value_input_option='USER_ENTERED')
                     
                    st.success("✅ PASO A COMPLETADO: Datos frescos cargados en Plantilla de forma instantánea.")
                    st.session_state['paso_a_listo'] = True
                except Exception as e:
                    st.error(f"🚨 Error en Paso A: {e}")
 
        st.markdown("---")
        st.markdown("### ⚡ PASO B: SINCRONIZADOR DE PRECIOS (ESTADO DEL ARSENAL)")
         
        if st.button("🔍 ESCANEAR ESTADO ACTUAL", use_container_width=True):
            with st.spinner("Escaneando el estado de la bóveda de precios..."):
                try:
                    gc = inicializar_cliente_gspread()
                    if gc is None:
                        st.error("🚨 Enlace satelital roto con Google Cloud.")
                        st.stop()
                        
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    sh = gc.open_by_url(url_boveda)
                    ws_conf = sh.worksheet("Configuración")
                     
                    data = ws_conf.get_all_values()
                    df_conf = pd.DataFrame(data[1:], columns=data[0])
                     
                    radar = df_conf.iloc[:, [8, 9, 10]].copy()
                    radar.columns = ['PRODUCTO', 'PRECIO_ACTUAL', 'PRECIO_SAP']
                     
                    # Filtro anti-fantasmas multi-formato
                    def es_fila_basura(val):
                        val_str = str(val).strip().upper()
                        if val_str in ["", "NAN", "NONE", "PRODUCTO"]: return True
                        try:
                            if float(val_str) == 0: return True
                        except ValueError:
                            pass
                        return False
                        
                    radar = radar[~radar['PRODUCTO'].apply(es_fila_basura)]
                     
                    radar['PRECIO_ACTUAL'] = radar['PRECIO_ACTUAL'].apply(extraer_numero)
                    radar['PRECIO_SAP'] = radar['PRECIO_SAP'].apply(extraer_numero)
                    radar['DIFERENCIA'] = (radar['PRECIO_SAP'] - radar['PRECIO_ACTUAL']).round(2)
                    radar['ESTADO'] = radar['DIFERENCIA'].apply(lambda x: "✅ OK" if x == 0 else "❌ DESFASE")
                    radar = radar.sort_values(by="ESTADO", ascending=False)
                     
                    # HUD OPERATIVO
                    total_insumos = len(radar)
                    insumos_ok = len(radar[radar['ESTADO'] == "✅ OK"])
                    insumos_fail = len(radar[radar['ESTADO'] == "❌ DESFASE"])
                     
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
                    def color_estado(val):
                        if val == "✅ OK": return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'
                        if val == "❌ DESFASE": return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center;'
                        return ''
 
                    st.dataframe(radar.style.map(color_estado, subset=['ESTADO']), use_container_width=True, hide_index=True)
                     
                    if insumos_fail == 0:
                        st.success("🟢 TODO EL SISTEMA ESTÁ EN NIVEL 'OK'. No se requieren ajustes operacionales.")
                    else:
                        st.warning("⚠️ SE DETECTARON DESFASES EN EL ARSENAL DE PRECIOS. Proceda a la inyección para nivelar los tableros.")
                        st.session_state['datos_para_sincronizar'] = True
 
                except Exception as e:
                    st.error(f"Error crítico al escanear los tableros: {e}")
 
        if st.session_state.get('datos_para_sincronizar'):
            st.markdown("---")
            if st.button("✅ APROBAR E INYECTAR PRECIOS (MODO SEGURO)", type="primary", use_container_width=True):
                with st.spinner("Inyectando quirúrgicamente Columna K en Columna J..."):
                    try:
                        gc = inicializar_cliente_gspread()
                        if gc is None:
                            st.error("🚨 Enlace satelital roto antes del volcado.")
                            st.stop()
                            
                        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                        ws_conf = sh.worksheet("Configuración")
                        data_full = ws_conf.get_all_values()
                        
                        valores_para_j = []
                        for fila in data_full[1:]:
                            valor_k = fila[10] if len(fila) > 10 else ""
                            valores_para_j.append([valor_k])
                         
                        if valores_para_j:
                            rango_destino = f"J2:J{len(valores_para_j) + 1}"
                            ws_conf.update(range_name=rango_destino, values=valores_para_j, value_input_option='USER_ENTERED')
                         
                        st.balloons()
                        st.success(f"🎯 INYECCIÓN EXITOSA. Se actualizaron {len(valores_para_j)} celdas en la columna J de forma segura.")
                        del st.session_state['datos_para_sincronizar']
                    except Exception as e:
                        st.error(f"🚨 FALLA EN LA INYECCIÓN EN CALIENTE: {e}")

if __name__ == "__main__":
    pass
