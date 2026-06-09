import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# ⚡ MOTOR DE CONEXIÓN PROPIO (Heredado del Módulo 8)
# =================================================================
@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread_propio():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

def cargar_datos_simulador(_procesar_fecha_pesada, _extraer_numero):
    gc = inicializar_cliente_gspread_propio()
    if not gc: return pd.DataFrame()
        
    try:
        url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
        sh = gc.open_by_url(url_maestra)
        ws = sh.worksheet("TABLA 1")
        datos_brutos = ws.get_all_values()
    except Exception:
        return pd.DataFrame()
    
    if not datos_brutos or len(datos_brutos) < 2: return pd.DataFrame()
        
    # 🧠 DETECCIÓN DINÁMICA DE ENCABEZADOS (Estilo Módulo 8)
    idx_headers = 4  
    for i in range(min(6, len(datos_brutos))):
        row_clean = [str(x).strip().upper() for x in datos_brutos[i]]
        if "Nº ORDEN" in row_clean or "FINCA" in row_clean or "PISTA" in row_clean:
            idx_headers = i
            break
            
    headers_limpios = []
    for h in datos_brutos[idx_headers]:
        h_str = str(h).strip().upper().replace("\n", " ")
        h_str = h_str.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
        h_str = h_str.replace("Ì", "I").replace("Ò", "O")
        headers_limpios.append(h_str)

    # MAPEO DE COORDENADAS
    idx_finca = headers_limpios.index("FINCA") if "FINCA" in headers_limpios else 2
    idx_ha = next((i for i, h in enumerate(headers_limpios) if "FUMIG" in h or "HA" in h), 5)
    idx_fecha = headers_limpios.index("FECHA") if "FECHA" in headers_limpios else 7
    idx_horometro = next((i for i, h in enumerate(headers_limpios) if "ODOM" in h), 10)
    idx_modelo = headers_limpios.index("MODELO") if "MODELO" in headers_limpios else 17
    idx_pista = headers_limpios.index("PISTA") if "PISTA" in headers_limpios else 23
    idx_cobro = next((i for i, h in enumerate(headers_limpios) if "COSTO AVION ($/HA)" in h or "FACTURAR" in h), 19)

    filas_datos = datos_brutos[idx_headers + 1:]
    lista_procesada = []
    
    for r in filas_datos:
        max_indice = max(idx_finca, idx_ha, idx_fecha, idx_horometro, idx_modelo, idx_pista, idx_cobro)
        if len(r) <= max_indice:
            r = r + [""] * (max_indice - len(r) + 1)
            
        finca_val = str(r[idx_finca]).strip().upper()
        if not finca_val or finca_val in ["NONE", "NAN", "FINCA", ""]: continue
            
        ha_netas = _extraer_numero(r[idx_ha])
        if ha_netas <= 0: continue # Solo pasamos vuelos reales
        
        horometro = _extraer_numero(r[idx_horometro])
        cobro_real_ha = _extraer_numero(r[idx_cobro])
        
        pista_raw = str(r[idx_pista]).strip().upper()
        pista_val = "PRINCIPAL" if not pista_raw or pista_raw in ["NONE", "NAN", ""] else pista_raw
        
        modelo_val = str(r[idx_modelo]).strip().upper()
        if not modelo_val or modelo_val in ["NONE", "NAN", ""]: continue
            
        f_str = str(r[idx_fecha]).strip()
        dt = _procesar_fecha_pesada(f_str)

        lista_procesada.append({
            "Fecha_DT": dt if dt else datetime.today(),
            "Finca": finca_val,
            "Pista": pista_val,
            "Equipo": modelo_val,
            "Hectareas": ha_netas,
            "Horometro": horometro,
            "CobroReal": cobro_real_ha
        })
        
    return pd.DataFrame(lista_procesada)

# =================================================================
# 👑 INTERFAZ DEL SIMULADOR
# =================================================================
def ejecutar(procesar_fecha_pesada, extraer_numero):
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis Inteligente de Lucro Cesante y Rendimiento Matemático de Flota.")

    with st.spinner("📥 Sincronizando datos con el Motor Táctico del Módulo 8..."):
        df_sim = cargar_datos_simulador(procesar_fecha_pesada, extraer_numero)

    if df_sim.empty:
        st.warning("📭 No hay registros matemáticamente válidos en la TABLA 1.")
        return

    # --- ELEMENTOS DE LOS FILTROS ---
    min_date = df_sim['Fecha_DT'].min().date()
    max_date = df_sim['Fecha_DT'].max().date()
    
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + sorted(df_sim["Finca"].dropna().unique().tolist())
    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + sorted(df_sim["Pista"].dropna().unique().tolist())
    opciones_avion = ["✈️ TODOS LOS EQUIPOS"] + sorted(df_sim["Equipo"].dropna().unique().tolist())
    lista_aviones_pura = sorted(df_sim["Equipo"].dropna().unique().tolist())

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL 
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario")
        f1, f2, f3, f4, f5, f6 = st.columns([1, 1, 1.2, 1, 1.2, 0.8])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        equipo_sel = f5.selectbox("✈️ Equipo Fijo", opciones_avion)
        multiplicador = f6.number_input("✖️ Mult.", value=1.112, format="%.3f")

        st.markdown("---")
        st.markdown("#### ✈️ Tarifas Dinámicas de Flota Real (Hora Avión / Ha Dron)")
        cols_av = st.columns(4)
        tarifas_aviones = {}
        for i, avion in enumerate(lista_aviones_pura):
            with cols_av[i % 4]:
                val_defecto = 84428.0 if "DRON" in avion or "DRONE" in avion else 4606562.0
                tarifas_aviones[avion] = st.number_input(f"💰 {avion}", value=val_defecto, step=10000.0, key=f"av_{i}")

    # --- FILTRAR ---
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios en las fechas seleccionadas.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO CON LÓGICA DUAL REAL E IMPACTO POR HECTÁREA
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    
    def calcular_costo_ha(row):
        # Si el modelo es un Dron o el horómetro es 0, cobra tarifa plana por hectárea
        if "DRON" in row["Equipo"] or "DRONE" in row["Equipo"] or row["Horometro"] == 0:
            return row["Tarifa_Aplicada"] * multiplicador
        else:
            # Avión: Tarifa * Horas / Hectáreas
            return ((row["Tarifa_Aplicada"] * row["Horometro"]) / row["Hectareas"]) * multiplicador

    # Este es el costo IDEAL por hectárea
    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(calcular_costo_ha, axis=1)
    
    # Cálculos Totales
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    # Agrupación por Finca y Equipo
    df_agrupado = df_filtrado.groupby(["Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
        "Horometro": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()

    # 🌟 NUEVO: CÁLCULO DE DIFERENCIA POR HECTÁREA POST-AGRUPACIÓN
    # Dividimos los totales consolidados entre las hectáreas para sacar los promedios reales
    df_agrupado["Tarifa Real Prom/Ha"] = df_agrupado["Total Real Facturado"] / df_agrupado["Hectareas"]
    df_agrupado["Tarifa Ideal Prom/Ha"] = df_agrupado["Total Simulado Ideal"] / df_agrupado["Hectareas"]
    df_agrupado["Brecha por Ha"] = df_agrupado["Tarifa Ideal Prom/Ha"] - df_agrupado["Tarifa Real Prom/Ha"]

    # =================================================================
    # 📊 RENDERIZADO DEL DASHBOARD TÁCTICO
    # =================================================================
    st.markdown("---")
    st.markdown("### 💎 Impacto Financiero de la Operación")
    
    t_real = df_agrupado["Total Real Facturado"].sum()
    t_ideal = df_agrupado["Total Simulado Ideal"].sum()
    t_perdido = df_agrupado["Lucro Cesante"].sum()
    porcentaje_fuga = ((t_ideal / t_real) - 1) * 100 if t_real > 0 else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Cobro Real Registrado (Con Topes)", f"$ {t_real:,.0f}")
    m2.metric("Cobro Matemático Puro (Sin Topes)", f"$ {t_ideal:,.0f}")
    m3.metric("⚠️ Lucro Cesante (Brecha Total)", f"$ {t_perdido:,.0f}", delta=f"{porcentaje_fuga:.1f}% de fuga", delta_color="inverse")

    st.markdown("#### 📉 Comparativa Facturación Total por Finca")
    df_g_resumen = df_agrupado.groupby("Finca")[["Total Real Facturado", "Total Simulado Ideal"]].sum().reset_index()
    df_g = df_g_resumen.melt(id_vars="Finca", value_vars=["Total Real Facturado", "Total Simulado Ideal"], var_name="Escenario", value_name="Monto ($)")
    fig = px.bar(df_g, x="Finca", y="Monto ($)", color="Escenario", barmode="group",
                 color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Ideal": "#d4af37"})
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, legend=dict(yanchor="top", y=1.1, xanchor="left", x=0.01))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 📋 Análisis Detallado: Brecha por Hectárea y Total")
    # Preparamos la tabla para mostrar la economía unitaria
    df_mostrar = df_agrupado[["Finca", "Equipo", "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha", "Brecha por Ha", "Lucro Cesante"]].copy()
    
    # Formateo visual de la tabla
    df_mostrar["Hectareas"] = df_mostrar["Hectareas"].apply(lambda x: f"{x:,.1f}")
    df_mostrar["Tarifa Real Prom/Ha"] = df_mostrar["Tarifa Real Prom/Ha"].apply(lambda x: f"$ {x:,.0f}")
    df_mostrar["Tarifa Ideal Prom/Ha"] = df_mostrar["Tarifa Ideal Prom/Ha"].apply(lambda x: f"$ {x:,.0f}")
    df_mostrar["Brecha por Ha"] = df_mostrar["Brecha por Ha"].apply(lambda x: f"$ {x:,.0f}")
    df_mostrar["Lucro Cesante"] = df_mostrar["Lucro Cesante"].apply(lambda x: f"$ {x:,.0f}")
    
    st.dataframe(df_mostrar.sort_values(by=["Finca"]), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    pass
