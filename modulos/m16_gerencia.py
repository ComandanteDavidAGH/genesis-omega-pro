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
        if '.' in v and ',' not in v:
            partes = v.split('.')
            if len(partes) == 2 and len(partes[1]) == 3:
                v = v.replace('.', '')
        elif ',' in v:
            v = v.replace('.', '').replace(',', '.')
        return float(v)
    except:
        return 0.0

@st.cache_data(show_spinner=False, ttl=120)
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
            df['FECHA_RAW'] = df['FECHA'].apply(procesar_fecha_pesada)
            df['FECHA_DT'] = pd.to_datetime(df['FECHA_RAW'], errors='coerce')
            
            # 🔒 FILTRO OPERATIVO TOTAL: Solo Año Fiscal 2026
            df = df[df['FECHA_DT'].dt.year == 2026]
            
            def clasificar_tec(row):
                texto = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))} {str(row.get('PISTA',''))}".upper()
                if 'DRON' in texto or 'DR5' in texto: return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
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

    st.markdown("<h1 style='color: #1a365d; font-family: Arial Black; border-bottom: 3px solid #d4af37;'>📊 Panel de Eficiencia Gerencial (Año Fiscal 2026)</h1>", unsafe_allow_html=True)

    # --- 🛰️ PANEL DE CONTROL DE AUDITORÍA Y ESTABILIZACIÓN ---
    with st.container(border=True):
        st.markdown("#### 🛠️ Herramientas de Estabilización de Datos")
        c1, c2 = st.columns(2)
        
        # El interruptor que salvará la presentación frente a Gerencia
        activar_parche = c1.toggle("⚡ Activar Corrección Automática de Flota (Ignorar errores del pasado)", value=True, help="Si está activo, el sistema aproximará las tarifas de Dron que fueron corrompidas en el Excel al valor fijo oficial más cercano (71.280, 75.518, 84.428).")
        
        if c2.button("🔄 Forzar Sincronización Total con la Nube", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_raw = cargar_datos_gerenciales()
    
    if df_raw.empty:
        st.warning("⚠️ No se detectan registros operativos para el año 2026 en la TABLA 1.")
        return

    df_base = df_raw.copy()

    # 💥 APLICACIÓN DEL PARCHE ANTI-DECIMALES EN DRONES
    if activar_parche:
        def corregir_dron(row):
            val = row['COSTO_VUELO_HA']
            if row['TECNOLOGIA'] == 'DRONE' and val > 0:
                oficiales = [71280, 75518, 84428]
                return min(oficiales, key=lambda x: abs(x - val))
            return val
        df_base['COSTO_VUELO_HA'] = df_base.apply(corregir_dron, axis=1)

    # --- 🏗️ CONSTRUCCIÓN DE LAS PESTAÑAS ---
    tab_vuelo, tab_total, tab_auditoria = st.tabs(["✈️ EFICIENCIA PURA VUELO (Columna T)", "💰 ANALÍTICA COSTO TOTAL", "🔍 AUDITORÍA FILA POR FILA (Excel vs Sistema)"])

    def descargar_excel(df_comparativo):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_comparativo.to_excel(writer, index=False, sheet_name='Comparativo_2026')
        return output.getvalue()

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        return f"$ {val:,.0f}".replace(",", ".")

    # ==========================================
    # PESTAÑA 1: COSTO EXCLUSIVO VUELO (Métrica Pura)
    # ==========================================
    with tab_vuelo:
        st.success("🔬 Analizando estrictamente la Columna T: COSTO AVIÓN ($/ha) - Cero Insumos Químicos")
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
            fig.update_layout(title="Brecha Real de Tarifa Vuelo 2026 (Avión vs Dron)", barmode='group', plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(tickangle=-45))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("📌 No hay registros cruzados en 2026 para ambas tecnologías simultáneamente.")

    # ==========================================
    # PESTAÑA 2: COSTO TOTAL
    # ==========================================
    with tab_total:
        st.info("📊 Incluye: Químicos + Servicio de Vuelo + Margen de Distribución (Año 2026)")
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
    # PESTAÑA 3: LA PRUEBA REINA (Auditoría Fila por Finca)
    # ==========================================
    with tab_auditoria:
        st.markdown("#### 🔍 Historial de Registros Crudos en Google Sheets (2026)")
        st.caption("Esta pestaña te muestra de forma transparente la información cruda tal y como está guardada en tu Excel. Aquí podrás ver cuáles órdenes específicas arrastran los valores alterados de las pruebas del pasado.")
        
        df_audit_print = pd.DataFrame({
            "Nº OS": df_base["OS"],
            "FECHA EXCEL": df_base["FECHA"],
            "FINCA": df_base["FINCA"],
            "TECNOLOGÍA": df_base["TECNOLOGIA"],
            "EQUIPO": df_base["PILOTO"] + " / " + df_base["MODELO"],
            "VALOR EN TU EXCEL (Columna T)": df_base["COSTO_HA"],
            "PROCESADO": df_base["COSTO_VUELO_HA"].apply(formatear_pesos)
        })
        st.dataframe(df_audit_print, use_container_width=True, hide_index=True)
