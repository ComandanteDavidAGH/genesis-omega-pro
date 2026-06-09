import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

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
        df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1]) if len(t1) > idx_t1 else pd.DataFrame()
        return df_t1
    except:
        return pd.DataFrame()

def limpiar_numero(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).replace("$", "").replace(" ", "").replace(",", "").replace("COP", "").strip()
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

    # --- AUTO-DETECCIÓN DE COLUMNAS (IA BÁSICA) ---
    cols = df_base.columns.tolist()
    col_fecha = next((c for c in cols if 'FECHA' in c.upper()), cols[0])
    col_finca = next((c for c in cols if 'FINCA' in c.upper() or 'CLIENTE' in c.upper()), cols[1])
    col_ha = next((c for c in cols if 'HECT' in c.upper() or 'HA' in c.upper() or 'CANT' in c.upper()), cols[2])
    col_horo = next((c for c in cols if 'HOROMETRO' in c.upper() or 'TIEMPO' in c.upper()), cols[3])
    col_vuelo = next((c for c in cols if 'VUELO' in c.upper() or 'TARIFA' in c.upper() or 'VALOR' in c.upper()), cols[-1])

    # --- PREPROCESAMIENTO INVISIBLE ---
    df_sim = df_base[[col_fecha, col_finca, col_ha, col_horo, col_vuelo]].copy()
    df_sim = df_sim[df_sim[col_finca].astype(str).str.strip() != ""] # Filtra vacíos
    
    # Formateo estricto de números y fechas
    df_sim["Hectáreas"] = df_sim[col_ha].apply(limpiar_numero)
    df_sim["Horómetro"] = df_sim[col_horo].apply(limpiar_numero)
    df_sim["Cobro Real"] = df_sim[col_vuelo].apply(limpiar_numero)
    df_sim['Fecha_DT'] = pd.to_datetime(df_sim[col_fecha], dayfirst=True, errors='coerce')
    
    # Eliminar datos nulos matemáticamente inservibles
    df_sim = df_sim[(df_sim["Hectáreas"] > 0) & (df_sim["Horómetro"] > 0)]

    # --- OBTENER RANGOS PARA FILTROS ---
    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    lista_fincas = sorted(df_sim[col_finca].dropna().unique().tolist())
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + lista_fincas

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL (Filtros)
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario y Parámetros")
        f1, f2, f3, f4 = st.columns(4)
        
        rango_fechas = f1.date_input("📅 Rango de Análisis", [min_date, max_date], min_value=min_date, max_value=max_date)
        finca_sel = f2.selectbox("📍 Selección de Finca", opciones_finca)
        tarifa_base_hora = f3.number_input("💰 Tarifa Avión (Hora Base)", value=4606562.0, step=10000.0)
        multiplicador = f4.number_input("✖️ Mult. General", value=1.112, format="%.3f")

    # --- APLICAR FILTROS DE INTERFAZ ---
    if len(rango_fechas) == 2:
        start_d, end_d = rango_fechas
        df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= start_d) & (df_sim['Fecha_DT'].dt.date <= end_d)].copy()
    else:
        df_filtrado = df_sim.copy()

    if finca_sel != "🌍 TODAS LAS FINCAS":
        df_filtrado = df_filtrado[df_filtrado[col_finca] == finca_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados en este rango de fechas o finca seleccionada.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO (Cálculo sin Topes)
    # =================================================================
    # Fórmula: ( (Tarifa Base * Horometro) / Ha ) * Ha * Mult = Tarifa Base * Horometro * Mult
    df_filtrado["Costo Simulado"] = ((tarifa_base_hora * df_filtrado["Horómetro"]) / df_filtrado["Hectáreas"]) * multiplicador
    
    # Totales facturados por fila
    df_filtrado["Total Real Facturado"] = df_filtrado["Cobro Real"] * df_filtrado["Hectáreas"]
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado"] * df_filtrado["Hectáreas"]
    
    # Brecha Financiera (Lo que se pierde por el tope)
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    # Agrupación Táctica
    df_agrupado = df_filtrado.groupby(col_finca).agg({
        "Hectáreas": "sum",
        "Horómetro": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()

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
    m1.metric("Cobro Real (Histórico con Topes)", f"$ {t_real:,.0f}")
    m2.metric("Cobro Matemático Puro (Sin Topes)", f"$ {t_ideal:,.0f}")
    m3.metric("⚠️ Dinero Dejado en la Mesa", f"$ {t_perdido:,.0f}", delta=f"{porcentaje_fuga:.1f}% de fuga", delta_color="inverse")

    c_grafico, c_tabla = st.columns([1.5, 1])

    with c_grafico:
        st.markdown("#### 📉 Comparativa Facturación por Finca")
        df_g = df_agrupado.melt(id_vars=col_finca, value_vars=["Total Real Facturado", "Total Simulado Ideal"], var_name="Escenario", value_name="Monto ($)")
        fig = px.bar(df_g, x=col_finca, y="Monto ($)", color="Escenario", barmode="group",
                     color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Ideal": "#d4af37"})
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, legend=dict(yanchor="top", y=1.1, xanchor="left", x=0.01))
        st.plotly_chart(fig, use_container_width=True)

    with c_tabla:
        st.markdown("#### 📋 Reporte de Fugas")
        df_mostrar = df_agrupado[[col_finca, "Hectáreas", "Lucro Cesante"]].copy()
        df_mostrar.columns = ["Finca", "Ha Voladas", "Dinero Perdido"]
        df_mostrar["Dinero Perdido"] = df_mostrar["Dinero Perdido"].apply(lambda x: f"$ {x:,.0f}")
        df_mostrar["Ha Voladas"] = df_mostrar["Ha Voladas"].apply(lambda x: f"{x:,.1f}")
        st.dataframe(df_mostrar.sort_values(by="Dinero Perdido", ascending=False), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    ejecutar()
