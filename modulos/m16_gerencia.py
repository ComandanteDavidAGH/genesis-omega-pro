import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import gspread
import re
import io
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials

# 💥 CONEXIÓN A TU ARTILLERÍA NATIVA DE CONFIANZA
from modulos.utilidades import procesar_fecha_pesada

# =================================================================
# 🔌 CONEXIÓN Y MOTORES DE LIMPIEZA DE ALTA PRECISIÓN
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

def limpiar_tarifa_excel(val):
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip().replace("$", "").replace(" ", "")
    if not v or v in ['-', 'NAN', 'NONE']: return 0.0
    try:
        # Si el formato viene como "71.280", eliminamos el punto de miles para que Python lea 71280
        if '.' in v and ',' not in v:
            partes = v.split('.')
            if len(partes) == 2 and len(partes[1]) == 3:
                v = v.replace('.', '')
        elif ',' in v:
            v = v.replace('.', '').replace(',', '.')
        return float(v)
    except:
        return 0.0

@st.cache_data(show_spinner=False, ttl=600)
def cargar_datos_gerenciales():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    
    try:
        boveda_act = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        datos_brutos = boveda_act.worksheet("TABLA 1").get_all_values()
        
        if len(datos_brutos) > 5:
            columnas_t1 = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA"]
            filas_limpias = [r + [""]*(len(columnas_t1) - len(r)) for r in datos_brutos[5:]]
            df = pd.DataFrame([r[:len(columnas_t1)] for r in filas_limpias], columns=columnas_t1)
            
            df['FINCA'] = df['FINCA'].astype(str).str.strip().str.upper()
            
            # Sincronización perfecta con tus fechas reales del Excel
            df['FECHA_RAW'] = df['FECHA'].apply(procesar_fecha_pesada)
            df['FECHA_DT'] = pd.to_datetime(df['FECHA_RAW'], errors='coerce')
            
            # 🔒 CERROJO ABSOLUTO: Filtrado estricto en RAM solo para el año 2026
            df = df[df['FECHA_DT'].dt.year == 2026]
            
            def clasificar_tec(row):
                texto = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))} {str(row.get('PISTA',''))}".upper()
                if 'DRON' in texto or 'DR5' in texto: return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
            
            # Limpieza financiera sin riesgo de escalas alteradas
            df['COSTO_TOTAL_HA'] = df['VALOR_FACTURAR'].apply(limpiar_tarifa_excel)
            df['COSTO_VUELO_HA'] = df['COSTO_HA'].apply(limpiar_tarifa_excel)
            
            return df
        return pd.DataFrame()
    except: 
        return pd.DataFrame()

# =================================================================
# 👑 RENDERIZADO VISUAL
# =================================================================

def ejecutar():
    st.header("", anchor="inicio_modulo")

    st.markdown("<h1 style='color: #1a365d; font-family: Arial Black; border-bottom: 3px solid #d4af37;'>📊 Inteligencia Comparativa (Periodo Fijo 2026)</h1>", unsafe_allow_html=True)

    df_base = cargar_datos_gerenciales()
    
    if df_base.empty:
        st.warning("⚠️ No se detectan operaciones registradas correspondientes al año 2026.")
        return

    # Mensaje informativo ejecutivo para la junta
    st.info("📅 **FILTRO ACTIVO:** Análisis financiero cerrado exclusivamente para el **Año Fiscal 2026**.")

    tab_total, tab_vuelo = st.tabs(["💰 ANALÍTICA COSTO TOTAL", "✈️ EFICIENCIA PURA VUELO (Columna T)"])

    def descargar_excel(df_comparativo):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_comparativo.to_excel(writer, index=False, sheet_name='Comparativo_2026')
        return output.getvalue()

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        return f"$ {val:,.0f}".replace(",", ".")

    # ==========================================
    # PESTAÑA 1: COSTO TOTAL
    # ==========================================
    with tab_total:
        st.caption("Incluye: Químicos + Servicio de Vuelo + Margen de Distribución (Año 2026)")
        matriz_t = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_TOTAL_HA', aggfunc='mean').reset_index()
        if 'AVIÓN' not in matriz_t.columns: matriz_t['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_t.columns: matriz_t['DRONE'] = np.nan
        
        m_comp = matriz_t.dropna(subset=['AVIÓN', 'DRONE']).copy()
        
        if not m_comp.empty:
            m_comp['Diferencia ($)'] = m_comp['AVIÓN'] - m_comp['DRONE']
            m_comp['Eficiencia (%)'] = (m_comp['Diferencia ($)'] / m_comp['AVIÓN']) * 100
            
            df_print_t = m_comp.copy()
            df_print_t['AVIÓN'] = df_print_t['AVIÓN'].apply(formatear_pesos)
            df_print_t['DRONE'] = df_print_t['DRONE'].apply(formatear_pesos)
            df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($)'].apply(formatear_pesos)
            df_print_t['Eficiencia (%)'] = df_print_t['Eficiencia (%)'].map("{:+.1f}%".format)

            st.dataframe(df_print_t, use_container_width=True, hide_index=True)

            excel_data = descargar_excel(m_comp)
            st.download_button(label="📥 Descargar Reporte Total 2026 (Excel)", data=excel_data, file_name="Comparativo_Total_Fincas_2026.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("📌 No hay registros cruzados en 2026 para ambas tecnologías simultáneamente.")

    # ==========================================
    # PESTAÑA 2: COSTO EXCLUSIVO VUELO (Métrica Pura)
    # ==========================================
    with tab_vuelo:
        st.success("🔬 Métrica Pura de Aeronaves: Columna T (COSTO AVIÓN $/ha) sin insumos.")
        matriz_v = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_VUELO_HA', aggfunc='mean').reset_index()
        if 'AVIÓN' not in matriz_v.columns: matriz_v['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_v.columns: matriz_v['DRONE'] = np.nan
        
        m_comp_v = matriz_v.dropna(subset=['AVIÓN', 'DRONE']).copy()
        
        if not m_comp_v.empty:
            m_comp_v['Diferencia ($)'] = m_comp_v['AVIÓN'] - m_comp_v['DRONE']
            m_comp_v['Eficiencia (%)'] = (m_comp_v['Diferencia ($)'] / m_comp_v['AVIÓN']) * 100
            
            df_print_v = m_comp_v.copy()
            df_print_v['AVIÓN'] = df_print_v['AVIÓN'].apply(formatear_pesos)
            df_print_v['DRONE'] = df_print_v['DRONE'].apply(formatear_pesos)
            df_print_v['Diferencia ($)'] = df_print_v['Diferencia ($)'].apply(formatear_pesos)
            df_print_v['Eficiencia (%)'] = df_print_v['Eficiencia (%)'].map("{:+.1f}%".format)

            st.dataframe(df_print_v, use_container_width=True, hide_index=True)

            excel_data_v = descargar_excel(m_comp_v)
            st.download_button(label="📥 Descargar Reporte Vuelo 2026 (Excel)", data=excel_data_v, file_name="Comparativo_Tarifas_Vuelo_2026.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            m_comp_v['FINCA_CORTA'] = m_comp_v['FINCA'].str[:15]
            fig = go.Figure()
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['DRONE'], name='Dron', marker_color='#d4af37'))
            fig.update_layout(
                title="Brecha Real de Tarifa Vuelo 2026 (Avión vs Dron)", 
                barmode='group', 
                plot_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(tickangle=-45)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("📌 No hay registros cruzados en 2026 para ambas tecnologías simultáneamente.")
