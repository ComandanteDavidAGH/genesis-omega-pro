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

# 🛰️ ENLACES NATIVOS
from modulos.utilidades import procesar_fecha_pesada
from openpyxl.styles import PatternFill, Font, Alignment

# =================================================================
# 🔌 CONEXIÓN Y MOTORES DE LIMPIEZA FINANCIERA
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
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)): return res_nativo.date()
        if isinstance(res_nativo, date): return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except: return None

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
    except: return pd.DataFrame()

# =================================================================
# ⚙️ MOTOR EXCEL PROFESIONAL (ESTILO GERENCIAL)
# =================================================================

def generar_excel_maestro(df_total, df_vuelo):
    output = io.BytesIO()
    
    # Utilizamos openpyxl para darle formato de lujo
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_total.to_excel(writer, index=False, sheet_name='Facturación Total')
        df_vuelo.to_excel(writer, index=False, sheet_name='Tarifa Vuelo Pura')
        
        header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal='center', vertical='center')
        
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            
            # Ancho de columnas profesionales
            ws.column_dimensions['A'].width = 38  # FINCA
            ws.column_dimensions['B'].width = 18  # AVIÓN
            ws.column_dimensions['C'].width = 18  # DRONE
            ws.column_dimensions['D'].width = 20  # Diferencia
            ws.column_dimensions['E'].width = 16  # Eficiencia
            
            # Estilo Encabezados
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                
            # Formato financiero de celdas
            for row in range(2, ws.max_row + 1):
                ws[f'B{row}'].number_format = '"$"#,##0'
                ws[f'C{row}'].number_format = '"$"#,##0'
                ws[f'D{row}'].number_format = '"$"#,##0'
                ws[f'E{row}'].number_format = '0.0%' # Porcentaje nativo Excel
                
    return output.getvalue()

# =================================================================
# 👑 RENDERIZADO VISUAL EN PANTALLA
# =================================================================

def ejecutar():
    st.header("", anchor="inicio_modulo")
    st.markdown("<h1 style='color: #1a365d; font-family: Arial Black; border-bottom: 3px solid #d4af37;'>📊 Comparativo Financiero: Drone vs Avión</h1>", unsafe_allow_html=True)

    # --- 🛰️ FILTROS DE RANGO DE FECHAS ---
    with st.container(border=True):
        st.markdown("#### 📅 Parámetros del Reporte")
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

    # Filtrado estricto
    df_base = df_raw[(df_raw['FECHA_FILTRABLE'] >= fecha_inicio) & (df_raw['FECHA_FILTRABLE'] <= fecha_fin)].copy()
    if df_base.empty:
        st.error(f"❌ No se encontraron registros de vuelo para el rango seleccionado.")
        return

    # PREPARAR DATA EXCEL: COSTO TOTAL
    matriz_t = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_TOTAL_HA', aggfunc='max').reset_index()
    if 'AVIÓN' not in matriz_t.columns: matriz_t['AVIÓN'] = np.nan
    if 'DRONE' not in matriz_t.columns: matriz_t['DRONE'] = np.nan
    m_comp_t = matriz_t.dropna(subset=['AVIÓN', 'DRONE']).copy()
    if not m_comp_t.empty:
        m_comp_t['Diferencia ($)'] = m_comp_t['AVIÓN'] - m_comp_t['DRONE']
        # Dejamos la eficiencia en formato decimal (ej: 0.082) para que Excel lo vuelva %
        m_comp_t['Eficiencia (%)'] = m_comp_t['Diferencia ($)'] / m_comp_t['AVIÓN'] 

    # PREPARAR DATA EXCEL: VUELO PURO
    matriz_v = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_VUELO_HA', aggfunc='max').reset_index()
    if 'AVIÓN' not in matriz_v.columns: matriz_v['AVIÓN'] = np.nan
    if 'DRONE' not in matriz_v.columns: matriz_v['DRONE'] = np.nan
    m_comp_v = matriz_v.dropna(subset=['AVIÓN', 'DRONE']).copy()
    if not m_comp_v.empty:
        m_comp_v['Diferencia ($)'] = m_comp_v['AVIÓN'] - m_comp_v['DRONE']
        m_comp_v['Eficiencia (%)'] = m_comp_v['Diferencia ($)'] / m_comp_v['AVIÓN']

    # --- 💎 BOTÓN MAESTRO DE DESCARGA (Arriba para mayor elegancia) ---
    if not m_comp_t.empty and not m_comp_v.empty:
        excel_data = generar_excel_maestro(m_comp_t, m_comp_v)
        st.download_button(
            label="📥 DESCARGAR REPORTE GERENCIAL (EXCEL COMPLETO)", 
            data=excel_data, 
            file_name=f"Reporte_Eficiencia_{fecha_inicio}_al_{fecha_fin}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    tab_vuelo, tab_total = st.tabs(["✈️ TARIFA DE VUELO PURA (Columna T)", "💰 FACTURACIÓN TOTAL OPERACIÓN"])

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        return f"$ {val:,.0f}".replace(",", ".")

    # ==========================================================
    # PESTAÑA: TARIFA VUELO
    # ==========================================================
    with tab_vuelo:
        st.success("🔬 Eficiencia en tarifas de servicio por Finca (Cero Insumos).")
        if not m_comp_v.empty:
            df_print_v = m_comp_v.copy()
            df_print_v['AVIÓN'] = df_print_v['AVIÓN'].apply(formatear_pesos)
            df_print_v['DRONE'] = df_print_v['DRONE'].apply(formatear_pesos)
            df_print_v['Diferencia ($)'] = df_print_v['Diferencia ($)'].apply(formatear_pesos)
            df_print_v['Eficiencia (%)'] = (df_print_v['Eficiencia (%)'] * 100).map("{:+.1f}%".format)

            st.dataframe(df_print_v, use_container_width=True, hide_index=True)
            
            # Gráfica
            m_comp_v['FINCA_CORTA'] = m_comp_v['FINCA'].str[:15]
            fig = go.Figure()
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig.add_trace(go.Bar(x=m_comp_v['FINCA_CORTA'], y=m_comp_v['DRONE'], name='Dron', marker_color='#d4af37'))
            fig.update_layout(title="Brecha Real de Tarifa Vuelo (Avión vs Dron)", barmode='group', plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(tickangle=-45))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("📌 No hay datos cruzados en el rango.")

    # ==========================================================
    # PESTAÑA: COSTO TOTAL
    # ==========================================================
    with tab_total:
        st.info("📊 Impacto macro en presupuesto por Finca (Consolidado Completo)")
        if not m_comp_t.empty:
            df_print_t = m_comp_t.copy()
            df_print_t['AVIÓN'] = df_print_t['AVIÓN'].apply(formatear_pesos)
            df_print_t['DRONE'] = df_print_t['DRONE'].apply(formatear_pesos)
            df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($)'].apply(formatear_pesos)
            df_print_t['Eficiencia (%)'] = (df_print_t['Eficiencia (%)'] * 100).map("{:+.1f}%".format)

            st.dataframe(df_print_t, use_container_width=True, hide_index=True)
        else:
            st.warning("📌 No hay datos cruzados en el rango.")
