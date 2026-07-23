import streamlit as st
import pandas as pd
import numpy as np

# =================================================================
# ⚡ MOTOR DE CARGA Y PROCESAMIENTO ULTRARRÁPIDO EN CACHÉ
# =================================================================

@st.cache_data(show_spinner=False, ttl=600)
def cargar_inventario_supabase_cached():
    """Recupera el inventario persistente de Supabase en milisegundos."""
    if 'supabase' in st.session_state:
        try:
            supabase_client = st.session_state['supabase']
            respuesta = supabase_client.table("inventario_sap").select("*").execute()
            if respuesta.data and len(respuesta.data) > 0:
                return pd.DataFrame(respuesta.data)
        except Exception:
            pass
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def procesar_radar_logistico_cached(df_sabana):
    """
    Procesa, consolida y aplica reglas de seguridad sobre la Sábana SAP
    a velocidad luz usando vectorización pura.
    """
    if df_sabana.empty:
        return None, "VACIO", 0, 0, 0, pd.DataFrame()

    # 1. Mapeo estructural indexado
    col_cod = next((c for c in df_sabana.columns if str(c).strip() == 'Material'), None)
    col_pista = next((c for c in df_sabana.columns if str(c).strip() == 'Almacén'), None)
    col_saldo = next((c for c in df_sabana.columns if str(c).strip() == 'Libre utilización'), None)
    col_desc = next((c for c in df_sabana.columns if str(c).strip() == 'Descripción del material'), None)

    # Fallbacks automáticos
    if not col_cod: col_cod = next((c for c in df_sabana.columns if 'MATERIAL' in str(c).upper()), None)
    if not col_pista: col_pista = next((c for c in df_sabana.columns if 'ALMACEN' in str(c).upper() or 'LGORT' in str(c).upper()), None)
    if not col_saldo: col_saldo = next((c for c in df_sabana.columns if 'LIBRE' in str(c).upper() or 'UTILIZACION' in str(c).upper() or 'LABST' in str(c).upper()), None)
    if not col_desc: col_desc = next((c for c in df_sabana.columns if 'DESC' in str(c).upper() or 'TEXTO' in str(c).upper()), None)

    if not col_cod or not col_pista or not col_saldo:
        return None, "ERROR_COLUMNAS", 0, 0, 0, pd.DataFrame()

    # 2. Copia y Limpieza Vectorial
    df_temp = df_sabana.copy()
    df_temp[col_saldo] = pd.to_numeric(df_temp[col_saldo].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    df_temp = df_temp[df_temp[col_saldo] > 0]
    
    if df_temp.empty:
        return None, "OK", 0, 0, 0, pd.DataFrame()

    codigos_limpios = df_temp[col_cod].astype(str).str.split('.').str[0].str.strip()
    
    if col_desc:
        df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | " + df_temp[col_desc].astype(str).str.strip().str.upper()
    else:
        df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | INSUMO QUÍMICO REGISTRADO"

    # 3. CONSOLIDACIÓN MAESTRA
    inventario_agrupado = df_temp.groupby([col_pista, 'PRODUCTO_RADAR'])[col_saldo].sum().reset_index()
    
    # 4. COMPILADOR DE REGLAS DE SEGURIDAD OPERATIVA (Vectorizado)
    pistas_series = inventario_agrupado[col_pista].astype(str).str.upper()
    productos_series = inventario_agrupado['PRODUCTO_RADAR'].astype(str).str.upper()
    
    es_pista_menor = pistas_series.str.contains("LUCI|TEHO", na=False)
    es_aceite = productos_series.str.contains("ACEITE|GRANEL|COMBUSTIBLE|DICAM", na=False)
    es_mancol = productos_series.str.contains("MANCOL|MANCOZEB|103680|104287", na=False)
    es_aditivo = productos_series.str.contains("ACONDICIONADOR|NATURAMIN|105980|108214|105296", na=False)
    
    condiciones = [
        es_aceite & es_pista_menor,
        es_aceite & ~es_pista_menor,
        es_mancol & es_pista_menor,
        es_mancol & ~es_pista_menor,
        es_aditivo
    ]
    
    valores_limite = [1000, 30280, 1000, 2500, 30]
    regles_texto = [
        "1.000 L (Aceite - Pista Menor)",
        "30,280 L (Aceite - Pista Principal)",
        "1,000 L (Mancol - Pista Menor)",
        "2,500 L (Mancol - Pista Principal)",
        "30 L/Kg (Aditivo de Alta Rotación)"
    ]
    
    inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD'] = np.select(condiciones, valores_limite, default=100)
    inventario_agrupado['📋 REGLA APLICADA'] = np.select(condiciones, regles_texto, default="100 L/Kg (Estándar Global)")
    
    # 5. Filtrar stock crítico
    df_alertas = inventario_agrupado[inventario_agrupado[col_saldo] < inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD']].copy()
    
    df_alertas = df_alertas.rename(columns={
        col_pista: "📍 PISTA / ALMACÉN",
        'PRODUCTO_RADAR': "🧪 CÓDIGO | NOMBRE DEL PRODUCTO",
        col_saldo: "⚠️ SALDO ACTUAL"
    })
    
    columnas_finales = ["📍 PISTA / ALMACÉN", "🧪 CÓDIGO | NOMBRE DEL PRODUCTO", "⚠️ SALDO ACTUAL", "🛡️ LÍMITE DE SEGURIDAD", "📋 REGLA APLICADA"]
    
    if not df_alertas.empty:
        df_alertas_render = df_alertas[columnas_finales].sort_values(by="📍 PISTA / ALMACÉN")
    else:
        df_alertas_render = pd.DataFrame(columns=columnas_finales)
    
    total_almacenes = inventario_agrupado[col_pista].nunique()
    total_insumos = inventario_agrupado['PRODUCTO_RADAR'].nunique()
    conteo_alertas = len(df_alertas_render)

    return df_alertas_render, "EXITO", total_almacenes, total_insumos, conteo_alertas, df_alertas_render

# =================================================================
# 👑 RENDERIZADO VISUAL DEL CENTRO DE MANDO
# =================================================================

def renderizar():
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
        box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important;
        overflow: hidden !important;
    }
    
    .hud-mando {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-left: 5px solid #0d1b2a;
        padding: 12px 20px;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #dee2e6;
    }
    .hud-mando-item { text-align: center; }
    .hud-mando-title { font-size: 11px; color: #6c757d; font-family: 'Arial Black', sans-serif; text-transform: uppercase; margin: 0; }
    .hud-mando-value { font-size: 20px; color: #0d1b2a; font-weight: 900; margin: 0; }
    .hud-mando-alert { color: #cc0000; font-family: 'Arial Black', sans-serif; }
    .hud-mando-ok { color: #00994c; font-family: 'Arial Black', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>🏠 Centro de Mando y Control</h1>", unsafe_allow_html=True)
    
    st.info("📡 **Radar Principal:** Monitoreo activo de sistemas, escuadrones y logística aérea.")
    st.markdown(f"### Bienvenido al Cuartel General, **{st.session_state.get('usuario_nombre', 'Comandante')}**.")
    st.write("El sistema Génesis Omega Pro se encuentra en línea y operando bajo parámetros óptimos. Seleccione un hangar en el menú lateral para iniciar operaciones.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 🚨 Radar Logístico: Alerta Temprana de Inventarios")
    
    df_sabana = st.session_state.get('df_sabana', pd.DataFrame())
    
    # ⚡ Recuperación en Caché desde Supabase
    if df_sabana.empty:
        df_sabana = cargar_inventario_supabase_cached()
        if not df_sabana.empty:
            st.session_state['df_sabana'] = df_sabana

    if df_sabana.empty:
        st.warning("⚠️ **Radar en Modo Espera:** El sistema no detecta un inventario activo en la memoria ni en la nube. Para encender el radar, por favor cargue la **Sábana SAP** actualizada en el **📥 Módulo 2 (Carga Facturación)**.")
    else:
        # ⚡ Ejecución Instantánea
        df_alertas_render, estado, total_almacenes, total_insumos, conteo_alertas, df_alertas_raw = procesar_radar_logistico_cached(df_sabana)

        if estado == "ERROR_COLUMNAS":
            st.error("❌ Error de Radar: No se pudieron mapear las columnas. Verifique que el archivo corresponda a la Sábana Estándar o estructura de Supabase.")
        else:
            clase_alerta = "hud-mando-value hud-mando-alert" if conteo_alertas > 0 else "hud-mando-value hud-mando-ok"
            texto_alerta = f"{conteo_alertas} Alertas" if conteo_alertas > 0 else "0 Críticos"
            
            st.markdown(f"""
            <div class="hud-mando">
                <div class="hud-mando-item">
                    <p class="hud-mando-title">Pistas / Almacenes Activos</p>
                    <p class="hud-mando-value">🛰️ {total_almacenes}</p>
                </div>
                <div class="hud-mando-item">
                    <p class="hud-mando-title">Insumos Consolidados Únicos</p>
                    <p class="hud-mando-value">🧪 {total_insumos}</p>
                </div>
                <div class="hud-mando-item">
                    <p class="hud-mando-title">Estado de Carga</p>
                    <p class="{clase_alerta}">{texto_alerta}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            if conteo_alertas > 0:
                st.error("🚨 **¡ALERTA ROJA! MÁRGENES OPERATIVOS CRÍTICOS DETECTADOS:**")
                
                df_render_display = df_alertas_render.copy()
                df_render_display["⚠️ SALDO ACTUAL"] = df_render_display["⚠️ SALDO ACTUAL"].apply(lambda x: f"{x:,.1f}".replace(",", "."))
                df_render_display["🛡️ LÍMITE DE SEGURIDAD"] = df_render_display["🛡️ LÍMITE DE SEGURIDAD"].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                
                def pintar_rojo_elegante(val):
                    return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold; border-bottom: 1px solid #dee2e6;'] * len(val)
                
                st.dataframe(df_render_display.style.apply(pintar_rojo_elegante, axis=1), use_container_width=True, hide_index=True)
            else:
                st.success("✅ **INVENTARIO ÓPTIMO:** Todos los insumos químicos y energéticos en la totalidad de las pistas se encuentran por encima de los márgenes de seguridad establecidos. Operación aérea asegurada.")
