import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 CONEXIÓN A BÓVEDA DE DATOS (Reciclado de tu Arquitectura)
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
    if pd.isna(val) or val == "": return 0.0
    try:
        # Limpia signos de dólar, espacios y convierte comas a puntos si es necesario
        texto = str(val).replace("$", "").replace(" ", "").replace(",", "")
        return float(texto)
    except:
        return 0.0

# =================================================================
# 🚁 MOTOR DEL SIMULADOR SIN TOPES
# =================================================================
def ejecutar():
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #00ff00; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador de Rendimiento Agroaéreo (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis de Lucro Cesante: Comparación de precios topados vs matemáticos reales.")

    with st.spinner("📥 Descargando historial de operaciones desde TABLA 1..."):
        df_base = extraer_tabla_1_historica()

    if df_base.empty:
        st.error("🚨 No se pudo conectar a TABLA 1 o está vacía.")
        return

    # 1. PARAMETRIZACIÓN DEL SIMULADOR
    with st.expander("⚙️ Calibrar Variables del Simulador", expanded=True):
        st.markdown("Ajuste los valores base para la fórmula matemática pura:")
        # Fórmula en LaTeX para profesionalismo
        st.latex(r"\text{Costo Puro} = \left( \frac{\text{Tarifa Base Hora} \times \text{Horómetro}}{\text{Hectáreas}} \right) \times \text{Multiplicador}")
        
        c1, c2 = st.columns(2)
        tarifa_base_hora = c1.number_input("Tarifa Base de Avión (Hora)", value=4606562.0, step=10000.0)
        multiplicador_productor = c2.number_input("Multiplicador Estándar (Ej: 1.112)", value=1.112, format="%.3f")

    st.markdown("---")
    st.markdown("### 🗺️ Mapeo de Columnas (Inteligencia de Datos)")
    st.info("Seleccione qué columnas de TABLA 1 contienen la información histórica a analizar.")
    
    col_cols = st.columns(4)
    columnas_disp = df_base.columns.tolist()
    
    # Auto-detección de columnas
    sug_finca = next((c for c in columnas_disp if "FINCA" in c.upper()), columnas_disp[0])
    sug_ha = next((c for c in columnas_disp if "HECT" in c.upper() or "HA" in c.upper()), columnas_disp[0])
    sug_horo = next((c for c in columnas_disp if "HOROMETRO" in c.upper() or "TIEMPO" in c.upper()), columnas_disp[0])
    sug_valor_real = next((c for c in columnas_disp if "VUELO" in c.upper() or "TARIFA" in c.upper() or "VALOR" in c.upper()), columnas_disp[0])

    col_finca = col_cols[0].selectbox("Columna: Finca", columnas_disp, index=columnas_disp.index(sug_finca))
    col_ha = col_cols[1].selectbox("Columna: Hectáreas", columnas_disp, index=columnas_disp.index(sug_ha))
    col_horo = col_cols[2].selectbox("Columna: Horómetro", columnas_disp, index=columnas_disp.index(sug_horo))
    col_valor_real = col_cols[3].selectbox("Columna: Cobro Real Vuelo", columnas_disp, index=columnas_disp.index(sug_valor_real))

    if st.button("🚀 EJECUTAR SIMULACIÓN MASIVA", type="primary"):
        # Preparar datos limpios
        df_sim = df_base[[col_finca, col_ha, col_horo, col_valor_real]].copy()
        df_sim = df_sim[df_sim[col_finca].astype(str).str.strip() != ""] # Quitar vacíos
        
        # Limpieza matemática
        df_sim["Hectáreas"] = df_sim[col_ha].apply(limpiar_numero)
        df_sim["Horómetro"] = df_sim[col_horo].apply(limpiar_numero)
        df_sim["Cobro Real (Histórico)"] = df_sim[col_valor_real].apply(limpiar_numero)

        # Filtrar datos inválidos (cero hectáreas o cero horómetro)
        df_sim = df_sim[(df_sim["Hectáreas"] > 0) & (df_sim["Horómetro"] > 0)]

        # 🧠 EL MOTOR SIN TOPES: Cálculo puro
        df_sim["Costo Simulado (Sin Topes)"] = ((tarifa_base_hora * df_sim["Horómetro"]) / df_sim["Hectáreas"]) * multiplicador_productor
        
        # Como es costo total facturado en la finca (multiplicado por hectareas):
        df_sim["Total Real Facturado"] = df_sim["Cobro Real (Histórico)"] * df_sim["Hectáreas"]
        df_sim["Total Simulado Facturable"] = df_sim["Costo Simulado (Sin Topes)"] * df_sim["Hectáreas"]
        
        # Diferencia (Lucro Cesante)
        df_sim["Diferencia (Dinero Perdido)"] = df_sim["Total Simulado Facturable"] - df_sim["Total Real Facturado"]

        # Agrupar por Finca para el Dashboard
        df_agrupado = df_sim.groupby(col_finca).agg({
            "Hectáreas": "sum",
            "Horómetro": "sum",
            "Total Real Facturado": "sum",
            "Total Simulado Facturable": "sum",
            "Diferencia (Dinero Perdido)": "sum"
        }).reset_index()

        # =================================================================
        # 📊 RENDERIZADO DEL DASHBOARD FINANCIERO
        # =================================================================
        st.markdown("---")
        st.markdown("### 💰 Impacto Financiero Global")
        total_real = df_agrupado["Total Real Facturado"].sum()
        total_sim = df_agrupado["Total Simulado Facturable"].sum()
        lucro_cesante = df_agrupado["Diferencia (Dinero Perdido)"].sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Facturado (Con Topes)", f"$ {total_real:,.0f}")
        m2.metric("Total Ideal (Sin Topes)", f"$ {total_sim:,.0f}")
        m3.metric("💸 Lucro Cesante (Brecha)", f"$ {lucro_cesante:,.0f}", delta=f"{((total_sim/total_real)-1)*100:.1f}%" if total_real > 0 else "0%")

        st.markdown("#### 📉 Comparativa de Facturación por Finca")
        
        # Preparar datos para gráfico de barras doble
        df_grafico = df_agrupado.melt(id_vars=col_finca, value_vars=["Total Real Facturado", "Total Simulado Facturable"], var_name="Escenario", value_name="Monto ($)")
        
        fig = px.bar(df_grafico, x=col_finca, y="Monto ($)", color="Escenario", barmode="group",
                     color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Facturable": "#00ff00"})
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=400)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 📋 Matriz Detallada de Rendimiento y Brechas")
        # Formatear matriz para visualización
        df_mostrar = df_agrupado.copy()
        df_mostrar.columns = ["Finca", "Ha Totales", "Horas Voladas", "Cobrado (Real)", "Debería Cobrarse", "Dinero Perdido por Topes"]
        for c in ["Cobrado (Real)", "Debería Cobrarse", "Dinero Perdido por Topes"]:
            df_mostrar[c] = df_mostrar[c].apply(lambda x: f"$ {x:,.0f}")
        
        st.dataframe(df_mostrar.sort_values(by="Dinero Perdido por Topes", ascending=False), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    ejecutar()
