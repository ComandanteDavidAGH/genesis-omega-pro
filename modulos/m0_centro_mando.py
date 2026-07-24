import streamlit as st
import pandas as pd
import numpy as np
import gspread

def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=600)
def cargar_inventario_supabase_cached():
    if 'supabase' in st.session_state and st.session_state['supabase'] is not None:
        try:
            supabase_client = st.session_state['supabase']
            respuesta = supabase_client.table("inventario_sap").select("*").execute()
            if respuesta.data and len(respuesta.data) > 0:
                return pd.DataFrame(respuesta.data)
        except Exception: pass
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def procesar_radar_logistico_cached(df_sabana):
    if df_sabana.empty: return None, "VACIO", 0, 0, 0, pd.DataFrame()
    col_cod = next((c for c in df_sabana.columns if str(c).strip() == 'Material'), None)
    col_pista = next((c for c in df_sabana.columns if str(c).strip() == 'Almacén'), None)
    col_saldo = next((c for c in df_sabana.columns if str(c).strip() == 'Libre utilización'), None)
    col_desc = next((c for c in df_sabana.columns if str(c).strip() == 'Descripción del material'), None)

    if not col_cod: col_cod = next((c for c in df_sabana.columns if 'MATERIAL' in str(c).upper()), None)
    if not col_pista: col_pista = next((c for c in df_sabana.columns if 'ALMACEN' in str(c).upper() or 'LGORT' in str(c).upper()), None)
    if not col_saldo: col_saldo = next((c for c in df_sabana.columns if 'LIBRE' in str(c).upper() or 'UTILIZACION' in str(c).upper() or 'LABST' in str(c).upper()), None)
    if not col_desc: col_desc = next((c for c in df_sabana.columns if 'DESC' in str(c).upper() or 'TEXTO' in str(c).upper()), None)

    if not col_cod or not col_pista or not col_saldo: return None, "ERROR_COLUMNAS", 0, 0, 0, pd.DataFrame()

    df_temp = df_sabana.copy()
    df_temp[col_saldo] = pd.to_numeric(df_temp[col_saldo].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    df_temp = df_temp[df_temp[col_saldo] > 0]
    if df_temp.empty: return None, "OK", 0, 0, 0, pd.DataFrame()

    codigos_limpios = df_temp[col_cod].astype(str).str.split('.').str[0].str.strip()
    df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | " + df_temp[col_desc].astype(str).str.strip().str.upper() if col_desc else codigos_limpios + " | INSUMO QUÍMICO REGISTRADO"

    inventario_agrupado = df_temp.groupby([col_pista, 'PRODUCTO_RADAR'])[col_saldo].sum().reset_index()
    
    condiciones = [
        inventario_agrupado[col_pista].astype(str).str.upper().str.contains("LUCI|TEHO", na=False) & inventario_agrupado['PRODUCTO_RADAR'].astype(str).str.upper().str.contains("ACEITE|GRANEL|COMBUSTIBLE", na=False)
    ]
    inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD'] = np.select(condiciones, [1000], default=100)
    inventario_agrupado['📋 REGLA APLICADA'] = np.select(condiciones, ["Regla Activa"], default="Estándar")
    
    df_alertas = inventario_agrupado[inventario_agrupado[col_saldo] < inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD']].copy()
    df_alertas = df_alertas.rename(columns={col_pista: "📍 PISTA / ALMACÉN", 'PRODUCTO_RADAR': "🧪 CÓDIGO | NOMBRE DEL PRODUCTO", col_saldo: "⚠️ SALDO ACTUAL"})
    
    columnas_finales = ["📍 PISTA / ALMACÉN", "🧪 CÓDIGO | NOMBRE DEL PRODUCTO", "⚠️ SALDO ACTUAL", "🛡️ LÍMITE DE SEGURIDAD", "📋 REGLA APLICADA"]
    df_alertas_render = df_alertas[columnas_finales].sort_values(by="📍 PISTA / ALMACÉN") if not df_alertas.empty else pd.DataFrame(columns=columnas_finales)
    
    return df_alertas_render, "EXITO", inventario_agrupado[col_pista].nunique(), inventario_agrupado['PRODUCTO_RADAR'].nunique(), len(df_alertas_render), df_alertas

# =================================================================
# 🎣 LANZAR CEBO / SONDA DE DIAGNÓSTICO
# =================================================================
def lanzar_sonda_diagnostico():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        st.error("🚨 Sin conexión activa a Supabase.")
        return

    try:
        supabase = st.session_state['supabase']
        with st.status("🎣 Lanzando Cebo a Supabase...", expanded=True) as status:
            
            st.write("1. Inyectando misión falsa (`_CEBO_`)...")
            supabase.table("TABLA_1").insert({"Nº ORDEN": "_CEBO_"}).execute()
            
            st.write("2. Leyendo cómo PostgreSQL nombra las columnas...")
            res = supabase.table("TABLA_1").select("*").eq("Nº ORDEN", "_CEBO_").execute()
            
            if res.data and len(res.data) > 0:
                columnas_reales = list(res.data[0].keys())
                
                st.write("3. Destruyendo cebo para limpiar base de datos...")
                supabase.table("TABLA_1").delete().eq("Nº ORDEN", "_CEBO_").execute()
                
                status.update(label="🎯 CEBO CAPTURADO CON ÉXITO", state="complete")
                
                st.success("✅ **Comandante, aquí están las entrañas exactas de su base de datos. Por favor copie todo el texto del recuadro oscuro y envíemelo:**")
                st.code(repr(columnas_reales), language="python")
            else:
                st.error("🚨 El cebo se inyectó pero no se pudo leer.")

    except Exception as e:
        st.error(f"🚨 Falla en el cebo: {e}")

# =================================================================
# 👑 RENDERIZADO VISUAL DEL CENTRO DE MANDO
# =================================================================
def renderizar():
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important; overflow: hidden !important; }
    .hud-mando { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border-left: 5px solid #0d1b2a; padding: 12px 20px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; box-shadow: 2px 2px 8px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #dee2e6; }
    .hud-mando-item { text-align: center; }
    .hud-mando-title { font-size: 11px; color: #6c757d; font-family: 'Arial Black', sans-serif; text-transform: uppercase; margin: 0; }
    .hud-mando-value { font-size: 20px; color: #0d1b2a; font-weight: 900; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>🏠 Centro de Mando y Control</h1>", unsafe_allow_html=True)
    st.info("📡 **Radar Principal:** Monitoreo activo de sistemas, escuadrones y logística aérea.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 🎣 Panel de Diagnóstico Táctico (EL CEBO)")
    st.warning("⚠️ **ATENCIÓN COMANDANTE:** Haga clic en el botón de abajo para extraer el código genético de su Supabase y acabar con el error de las columnas.")
    
    if st.button("🎣 LANZAR CEBO DE DIAGNÓSTICO", type="secondary", use_container_width=True):
        lanzar_sonda_diagnostico()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 🚨 Radar Logístico")
    
    df_sabana = st.session_state.get('df_sabana', pd.DataFrame())
    if df_sabana.empty:
        df_sabana = cargar_inventario_supabase_cached()
        if not df_sabana.empty: st.session_state['df_sabana'] = df_sabana

    if df_sabana.empty:
        st.warning("⚠️ El sistema no detecta un inventario activo. Cargue la Sábana SAP en el Módulo 2.")
    else:
        df_alertas_render, estado, total_almacenes, total_insumos, conteo_alertas, _ = procesar_radar_logistico_cached(df_sabana)
        if estado == "EXITO":
            st.markdown(f"""
            <div class="hud-mando">
                <div class="hud-mando-item"><p class="hud-mando-title">Pistas Activas</p><p class="hud-mando-value">🛰️ {total_almacenes}</p></div>
                <div class="hud-mando-item"><p class="hud-mando-title">Insumos Únicos</p><p class="hud-mando-value">🧪 {total_insumos}</p></div>
                <div class="hud-mando-item"><p class="hud-mando-title">Alertas</p><p class="hud-mando-value" style="color:{'#cc0000' if conteo_alertas > 0 else '#00994c'};">{conteo_alertas}</p></div>
            </div>
            """, unsafe_allow_html=True)
            if conteo_alertas > 0: st.dataframe(df_alertas_render, hide_index=True)
