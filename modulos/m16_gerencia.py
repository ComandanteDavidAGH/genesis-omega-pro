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

def limpiar_dinero(val):
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip()
    if not v or v == '-': return 0.0
    v = re.sub(r'[^\d\.,\-]', '', v)
    if not v: return 0.0
    try:
        if '.' in v and ',' in v:
            if v.rfind(',') > v.rfind('.'): 
                v = v.replace('.', '').replace(',', '.')
            else:
                v = v.replace(',', '')
        elif ',' in v: 
            v = v.replace(',', '.')
        
        num = float(v) if v else 0.0
        # 💥 PARCHE DE ESCALA FINANCIERO: Si Python lo lee como decimal entre 5 y 2500, lo escala a miles reales
        if 5 < num < 2500: 
            num = num * 1000
        return num
    except:
        return 0.0

def limpiar_area(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip()
        if not v: return 0.0
        v = v.replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1:
            partes = v.rsplit('.', 1)
            v = partes[0].replace('.', '') + '.' + partes[1]
        return float(v) if v else 0.0
    except: return 0.0

def procesar_fecha_gerencia(val):
    try:
        if isinstance(val, datetime): return val.date()
        s = str(val).strip().lower()
        # Limpieza de nombres de días si vienen en texto largo desde Excel
        if ',' in s: s = s.split(',')[1].strip()
        
        meses = {'enero':1, 'febrero':2, 'marzo':3, 'abril':4, 'mayo':5, 'junio':6, 'julio':7, 'agosto':8, 'septiembre':9, 'octubre':10, 'noviembre':11, 'diciembre':12}
        match = re.search(r'([a-z]+)\s+(\d{1,2}),\s+(\d{4})', s)
        if match:
            mes_str, dia_str, anio_str = match.groups()
            if mes_str in meses: return date(int(anio_str), meses[mes_str], int(dia_str))
            
        match2 = re.search(r'(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})', s)
        if match2:
            dia_str, mes_str, anio_str = match2.groups()
            if mes_str in meses: return date(int(anio_str), meses[mes_str], int(dia_str))

        return pd.to_datetime(s.split(" ")[0], dayfirst=True, errors='coerce').date()
    except: 
        return None

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
            df['FECHA_DT'] = df['FECHA'].apply(procesar_fecha_gerencia)
            
            def clasificar_tec(row):
                texto = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))}".upper()
                if 'DRON' in texto or 'DR5' in texto: return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
            
            # 💥 SEGURO CONTABLE: Inyección del motor limpiar_dinero para corregir las escalas automáticamente
            df['COSTO_TOTAL_HA'] = df['VALOR_FACTURAR'].apply(limpiar_dinero)
            df['COSTO_VUELO_HA'] = df['COSTO_HA'].apply(limpiar_dinero)
            
            return df.dropna(subset=['FECHA_DT'])
        return pd.DataFrame()
    except: 
        return pd.DataFrame()

# =================================================================
# 👑 RENDERIZADO VISUAL CON FORMATO LATINO
# =================================================================

def ejecutar():
    st.header("", anchor="inicio_modulo")

    st.markdown("<h1 style='color: #1a365d; font-family: Arial Black; border-bottom: 3px solid #d4af37;'>📊 Comparativo Drone vs Avión</h1>", unsafe_allow_html=True)

    # --- 🛰️ SELECTORES DE FECHA ---
    with st.container(border=True):
        c_f1, c_f2, c_f3 = st.columns([1, 1, 1])
        fecha_inicio = c_f1.date_input("📅 Desde:", value=date(2024, 1, 1))
        fecha_fin = c_f2.date_input("📅 Hasta:", value=datetime.now().date())
        
        if st.button("🔄 Forzar Recarga de Nube", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_raw = cargar_datos_gerenciales()
    
    if df_raw.empty:
        st.warning("⚠️ No se detectan datos en la base maestra.")
        return

    df_base = df_raw[(df_raw['FECHA_DT'] >= fecha_inicio) & (df_raw['FECHA_DT'] <= fecha_fin)].copy()

    if df_base.empty:
        st.error(f"❌ No se encontraron registros de vuelo entre el {fecha_inicio.strftime('%d/%m/%Y')} y el {fecha_fin.strftime('%d/%m/%Y')}")
        return

    tab_total, tab_vuelo = st.tabs(["💰 ANALÍTICA COSTO TOTAL", "✈️ EFICIENCIA PURA VUELO (Columna T)"])

    def descargar_excel(df_comparativo):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_comparativo.to_excel(writer, index=False, sheet_name='Comparativo')
        return output.getvalue()

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        return f"$ {val:,.0f}".replace(",", ".")

    # ==========================================
    # PESTAÑA 1: COSTO TOTAL
    # ==========================================
    with tab_total:
        st.info("📊 Incluye: Químicos + Servicio de Vuelo + Margen de Distribución")
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
            df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($を確認)'].apply(formatear_pesos) if 'Diferencia ($を確認)' in df_print_t else df_print_t['Diferencia ($)'].apply(formatear_pesos)
            df_print_t['Eficiencia (%)'] = df_print_t['Eficiencia (%)'].map("{:+.1f}%".format)

            st.dataframe(df_print_t, use_container_width=True, hide_index=True)

            excel_data = descargar_excel(m_comp)
            st.download_button(label="📥 Descargar Comparativo Total (Excel)", data=excel_data, file_name=f"Reporte_Total_{fecha_inicio}_al_{fecha_fin}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("📌 No hay fincas cruzadas que usaran ambas tecnologías en este rango de fechas.")

    # ==========================================
    # PESTAÑA 2: COSTO EXCLUSIVO VUELO
    # ==========================================
    with tab_vuelo:
        st.success("🔬 Analizando estrictamente la Columna T: COSTO AVIÓN ($/ha) - Cero Insumos")
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
            st.download_button(label="📥 Descargar Comparativo Vuelo (Excel)", data=excel_data_v, file_name=f"Reporte_Vuelo_{fecha_inicio}_al_{fecha_fin}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            m_comp_v['FINCA_CORTA'] = m_comp_v['FINCA'].str[:15]
            fig = go.Figure()
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['DRONE'], name='Dron', marker_color='#d4af37'))
            fig.update_layout(
                title="Brecha Real de Tarifa Vuelo (Avión vs Dron)", 
                barmode='group', 
                plot_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(tickangle=-45)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("📌 No hay fincas cruzadas que usaran ambas tecnologías en este rango de fechas.")
