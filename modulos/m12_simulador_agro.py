import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np # <-- Nueva librería para cálculos avanzados

# =================================================================
# 🔌 CONEXIÓN A BÓVEDA DE DATOS
# =================================================================
def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

@st.cache_data(show_spinner=False, ttl=600)
def extraer_tabla_1_historica():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        t1 = boveda.worksheet("TABLA 1").get_all_values()
        idx_t1 = 4
        for i in range(min(6, len(t1))):
            if "FINCA" in [str(x).upper() for x in t1[i]]:
                idx_t1 = i
                break
        
        if len(t1) > idx_t1:
            df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1])
            df_t1 = df_t1.loc[:, ~df_t1.columns.duplicated()]
            return df_t1
        return pd.DataFrame()
    except:
        return pd.DataFrame()

def limpiar_numero(val):
    if isinstance(val, pd.Series):
        val = val.iloc[0]
        
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).upper().replace("$", "").replace("COP", "").strip()
        
        # 🛡️ RESPETAR LA COMA DECIMAL (Ej: 3,5 -> 3.5)
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".") # Ej: 1.500,50
        elif "," in texto:
            texto = texto.replace(",", ".") # Ej: 3,5
            
        texto = texto.replace(" ", "")
        return float(texto)
    except:
        return 0.0

# =================================================================
# 🚁 MOTOR DEL SIMULADOR SIN TOPES
# =================================================================
def ejecutar():
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis Inteligente de Lucro Cesante y Rendimiento Matemático de Flota.")

    with st.spinner("📥 Sincronizando historial de operaciones (TABLA 1)..."):
        df_base = extraer_tabla_1_historica()

    if df_base.empty:
        st.error("🚨 Base de datos vacía o sin acceso a TABLA 1.")
        return

    # --- 🛡️ BUSCADOR INTELIGENTE Y BLINDADO DE COLUMNAS ---
    cols = df_base.columns.tolist()
    def buscar_col(keywords, excluir=[]):
        for c in cols:
            if c in excluir: continue
            c_up = str(c).upper()
            for kw in keywords:
                if kw in c_up: return c
        return None
    
    def fallback(excluir):
        for c in cols:
            if c not in excluir: return c
        return cols[0]

    col_fecha = buscar_col(['FECHA']) or fallback([])
    col_finca = buscar_col(['FINCA', 'CLIENTE'], excluir=[col_fecha]) or fallback([col_fecha])
    col_pista = buscar_col(['PISTA', 'ORIGEN', 'BASE'], excluir=[col_fecha, col_finca]) or fallback([col_fecha, col_finca])
    col_avion = buscar_col(['AVION', 'EQUIPO', 'AERONAVE', 'MAQUINA'], excluir=[col_fecha, col_finca, col_pista]) or fallback([col_fecha, col_finca, col_pista])
    # 🎯 PALABRAS CLAVE AJUSTADAS AL REPORTE
    col_ha = buscar_col(['ÁREA', 'AREA', 'HECT', 'CANT', 'HA'], excluir=[col_fecha, col_finca, col_pista, col_avion]) or fallback([col_fecha, col_finca, col_pista, col_avion])
    col_horo = buscar_col(['ODÒM', 'ODÓM', 'ODOM', 'HOROMETRO', 'TIEMPO', 'HORA'], excluir=[col_fecha, col_finca, col_pista, col_avion, col_ha]) or fallback([col_fecha, col_finca, col_pista, col_avion, col_ha])
    col_vuelo = buscar_col(['VUELO', 'TARIFA', 'VALOR', 'COBRO', 'TOTAL'], excluir=[col_fecha, col_finca, col_pista, col_avion, col_ha, col_horo]) or fallback([col_fecha, col_finca, col_pista, col_avion, col_ha, col_horo])

    # --- PREPROCESAMIENTO ABSOLUTO ---
    df_sim = df_base[[col_fecha, col_finca, col_pista, col_avion, col_ha, col_horo, col_vuelo]].copy()
    df_sim.columns = ["Fecha", "Finca", "Pista", "Equipo", "Hectareas", "Horometro", "CobroReal"]
    df_sim = df_sim[df_sim["Finca"].astype(str).str.strip() != ""] 
    
    df_sim["Hectareas"] = df_sim["Hectareas"].apply(limpiar_numero)
    df_sim["Horometro"] = df_sim["Horometro"].apply(limpiar_numero)
    df_sim["CobroReal"] = df_sim["CobroReal"].apply(limpiar_numero)
    df_sim['Fecha_DT'] = pd.to_datetime(df_sim["Fecha"], dayfirst=True, errors='coerce')
    
    # 🛡️ GUILLOTINA JUSTA: Solo eliminamos si las Hectáreas son 0. (Permitimos Horómetro 0 para los Drones)
    df_sim = df_sim[df_sim["Hectareas"] > 0]

    if df_sim.empty:
        st.warning("📭 No hay registros matemáticamente válidos (con hectáreas > 0) en la TABLA 1.")
        return

    # --- OBTENER RANGOS PARA FILTROS ---
    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime(2023, 1, 1).date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + sorted(df_sim["Finca"].dropna().unique().tolist())
    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + sorted(df_sim["Pista"].dropna().astype(str).unique().tolist())
    opciones_avion = ["✈️ TODOS LOS EQUIPOS"] + sorted(df_sim["Equipo"].dropna().astype(str).unique().tolist())
    
    lista_aviones_pura = sorted(df_sim["Equipo"].dropna().astype(str).unique().tolist())

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
        st.markdown("#### ✈️ Tarifas Dinámicas por Aeronave (Hora Avión / Ha Dron)")
        cols_av = st.columns(4)
        tarifas_aviones = {}
        for i, avion in enumerate(lista_aviones_pura):
            with cols_av[i % 4]:
                val_defecto = 84428.0 if "DRON" in avion.upper() else 4606562.0
                tarifas_aviones[avion] = st.number_input(f"💰 {avion}", value=val_defecto, step=10000.0, key=f"av_{i}")

    # --- APLICAR FILTROS ---
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos filtros.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO (Lógica Dual: Avión vs Dron)
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    
    # 🌟 NUEVA LÓGICA: Si el horómetro es mayor a 0, se asume que es Avión (costo = tarifa * horometro / ha). 
    # Si el horómetro es 0, se asume Drone (costo = tarifa por hectárea).
    def calcular_costo_ha(row):
        if row["Horometro"] > 0:
            return ((row["Tarifa_Aplicada"] * row["Horometro"]) / row["Hectareas"]) * multiplicador
        else:
            return row["Tarifa_Aplicada"] * multiplicador

    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(calcular_costo_ha, axis=1)
    
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    df_agrupado = df_filtrado.groupby(["Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
        "Horometro": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()

    # =================================================================
    # 📊 RENDERIZADO DEL DASHBOARD
    # =================================================================
    st.markdown("---")
    st.markdown("### 💎 Impacto Financiero de la Operación")
    
    t_real = df_agrupado["Total Real Facturado"].sum()
    t_ideal = df_agrupado["Total Simulado Ideal"].sum()
    t_perdido = df_agrupado["Lucro Cesante"].sum()
    porcentaje_fuga = ((t_ideal / t_real) - 1) * 100 if t_real > 0 else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Cobro Real (Con Topes)", f"$ {t_real:,.0f}")
    m2.metric("Cobro Matemático Puro", f"$ {t_ideal:,.0f}")
    m3.metric("⚠️ Dinero Dejado en la Mesa", f"$ {t_perdido:,.0f}", delta=f"{porcentaje_fuga:.1f}% de fuga", delta_color="inverse")

    c_grafico, c_tabla = st.columns([1.5, 1])

    with c_grafico:
        st.markdown("#### 📉 Comparativa Facturación por Finca")
        df_g_resumen = df_agrupado.groupby("Finca")[["Total Real Facturado", "Total Simulado Ideal"]].sum().reset_index()
        df_g = df_g_resumen.melt(id_vars="Finca", value_vars=["Total Real Facturado", "Total Simulado Ideal"], var_name="Escenario", value_name="Monto ($)")
        fig = px.bar(df_g, x="Finca", y="Monto ($)", color="Escenario", barmode="group",
                     color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Ideal": "#d4af37"})
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, legend=dict(yanchor="top", y=1.1, xanchor="left", x=0.01))
        st.plotly_chart(fig, use_container_width=True)

    with c_tabla:
        st.markdown("#### 📋 Reporte de Fugas")
        df_mostrar = df_agrupado[["Pista", "Finca", "Equipo", "Hectareas", "Lucro Cesante"]].copy()
        df_mostrar["Lucro Cesante"] = df_mostrar["Lucro Cesante"].apply(lambda x: f"$ {x:,.0f}")
        df_mostrar["Hectareas"] = df_mostrar["Hectareas"].apply(lambda x: f"{x:,.1f}")
        st.dataframe(df_mostrar.sort_values(by=["Lucro Cesante"], ascending=False), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    ejecutar()
