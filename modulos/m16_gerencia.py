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

# 🛰️ ENLACE DIRECTO A TU PROCESADOR DE FECHAS NATIVO
from modulos.utilidades import procesar_fecha_pesada

# =================================================================
# 🔌 CONEXIÓN Y MOTORES DE LIMPIEZA DE ALTA TRANSPARENCIA
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

def normalizar_a_fecha_pura(val):
    """Garantiza la conversión estricta a un objeto date puro para evitar fugas de memoria"""
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)):
            return res_nativo.date()
        if isinstance(res_nativo, date):
            return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except:
        return None

@st.cache_data(show_spinner=False, ttl=60)
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
            df['FECHA_RAW'] = df['FECHA']
            
            # 🔒 CANDADO 1: Conversión homogénea a fecha pura
            df['FECHA_FILTRABLE'] = df['FECHA'].apply(normalizar_a_fecha_pura)
            
            def clasificar_tec(row):
                texto = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))}".upper()
                if 'DRON' in texto or 'DR5' in texto: return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
            df['COSTO_TOTAL_HA'] = df['VALOR_FACTURAR'].apply(limpiar_tarifa_excel)
            df['COSTO_VUELO_HA'] = df['COSTO_HA'].apply(limpiar_tarifa_excel)
            
            return df.dropna(subset=['FECHA_FILTRABLE'])
        return pd.DataFrame()
    except: 
        return pd.DataFrame()

# =================================================================
# 👑 RENDERIZADO VISUAL CONTROLADO POR SELECTOR
# =================================================================

def ejecutar():
    st.header("", anchor="inicio_modulo")

    st.markdown("<h1 style='color: #1a365d; font-family: Arial Black; border-bottom: 3px solid #d4af37;'>📊 Inteligencia Comparativa Financiera</h1>", unsafe_allow_html=True)

    # --- 🛰️ SELECTORES DE FECHA INTERACTIVOS (RESTAURADOS) ---
    with st.container(border=True):
        st.markdown("#### 📅 Rango de Consulta Avanzada")
        c_f1, c_f2, c_f3 = st.columns([1, 1, 1])
        fecha_inicio = c_f1.date_input("Desde:", value=date(2026, 1, 1))
        fecha_fin = c_f2.date_input("Hasta:", value=date(2026, 12, 31))
        
        if c_f3.button("🔄 Sincronizar Google Drive", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_raw = cargar_datos_gerenciales()
    
    if df_raw.empty:
        st.warning("⚠️ No se detectan registros en la base maestra.")
        return

    # 🔒 CANDADO 2: Filtrado directo y estricto utilizando objetos tipo date puros
    df_base = df_raw[(df_raw['FECHA_FILTRABLE'] >= fecha_inicio) & (df_raw['FECHA_FILTRABLE'] <= fecha_fin)].copy()

    if df_base.empty:
        st.error(f"❌ No se encontraron registros de vuelo para el rango seleccionado: {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}")
        return

    tab_vuelo, tab_total = st.tabs(["✈️ TARIFA DE VUELO PURA (Columna T)", "💰 FACTURACIÓN TOTAL OPERACIÓN"])

    def descargar_excel(df_comparativo):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_comparativo.to_excel(writer, index=False, sheet_name='Reporte_Gerencial')
        return output.getvalue()

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        return f"$ {val:,.0f}".replace(",", ".")

    # ==========================================================
    # PESTAÑA 1: TARIFA DE VUELO PURA (Tu sección sagrada)
    # ==========================================================
    with tab_vuelo:
        st.success("🔬 Tarifas fijas del servicio extraídas directamente de la Columna T (Sin Químicos).")
        
        # 🔒 CANDADO 3: Desglose por OS y Fecha para destruir cualquier promedio flotante
        matriz_v = df_base.pivot_table(index=['FINCA', 'FECHA_RAW', 'OS'], columns='TECNOLOGIA', values='COSTO_VUELO_HA', aggfunc='first').reset_index()
        
        if 'AVIÓN' not in matriz_v.columns: matriz_v['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_v.columns: matriz_v['DRONE'] = np.nan
        
        m_comp_v = matriz_v.copy()
        m_comp_v['Diferencia ($)'] = m_comp_v['AVIÓN'].fillna(0) - m_comp_v['DRONE'].fillna(0)
        
        df_print_v = m_comp_v.copy()
        df_print_v['AVIÓN'] = df_print_v['AVIÓN'].apply(formatear_pesos)
        df_print_v['DRONE'] = df_print_v['DRONE'].apply(formatear_pesos)
        df_print_v['Diferencia ($)'] = df_print_v['Diferencia ($)'].apply(formatear_pesos)

        st.dataframe(df_print_v, use_container_width=True, hide_index=True)

        excel_data_v = descargar_excel(m_comp_v)
        st.download_button(label="📥 Descargar Reporte Tarifas (Excel)", data=excel_data_v, file_name=f"Tarifas_Vuelo_{fecha_inicio}_a_{fecha_fin}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ==========================================================
    # PESTAÑA 2: FACTURACIÓN TOTAL
    # ==========================================================
    with tab_total:
        st.info("📊 Consolidado completo por orden (Insumos + Aplicación)")
        matriz_t = df_base.pivot_table(index=['FINCA', 'FECHA_RAW', 'OS'], columns='TECNOLOGIA', values='COSTO_TOTAL_HA', aggfunc='first').reset_index()
        
        if 'AVIÓN' not in matriz_t.columns: matriz_t['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_t.columns: matriz_t['DRONE'] = np.nan
        
        m_comp_t = matriz_t.copy()
        m_comp_t['Diferencia ($)'] = m_comp_t['AVIÓN'].fillna(0) - m_comp_t['DRONE'].fillna(0)
        
        df_print_t = m_comp_t.copy()
        df_print_t['AVIÓN'] = df_print_t['AVIÓN'].apply(formatear_pesos)
        df_print_t['DRONE'] = df_print_t['DRONE'].apply(formatear_pesos)
        df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($)'].apply(formatear_pesos)

        st.dataframe(df_print_t, use_container_width=True, hide_index=True)

        excel_data = descargar_excel(m_comp_t)
        st.download_button(label="📥 Descargar Reporte Total (Excel)", data=excel_data, file_name=f"Facturacion_Total_{fecha_inicio}_a_{fecha_fin}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
