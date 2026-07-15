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

# 💥 TRANSLATOR PRO: Evita el colapso por puntos múltiples repetidos de SAP (Ej: 117.404.747)
def limpiar_tarifa_excel(val):
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip().replace("$", "").replace(" ", "").upper()
    if not v or v in ['-', 'NAN', 'NONE', '']: return 0.0
    
    s_clean = re.sub(r'[^\d\.,\-]', '', v)
    try:
        if '.' in s_clean and ',' in s_clean:
            if s_clean.rfind(',') > s_clean.rfind('.'): s_clean = s_clean.replace('.', '').replace(',', '.')
            else: s_clean = s_clean.replace(',', '')
        elif ',' in s_clean:
            if len(s_clean.split(',')[-1]) == 3: s_clean = s_clean.replace(',', '')
            else: s_clean = s_clean.replace(',', '.')
        elif '.' in s_clean:
            if s_clean.count('.') > 1: s_clean = s_clean.replace('.', '')
            elif len(s_clean.split('.')[-1]) == 3: s_clean = s_clean.replace('.', '')
        return float(s_clean) if s_clean else 0.0
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
            
            # 💥 ESCUDO ANTI-OUTLIERS (SANEAMIENTO DE ERRORES DE DIGITACIÓN EN SAP)
            # 1. Regla de Oro: Si digitan en cientos o decenas (ej: 65), se asume miles (65.000)
            df['COSTO_VUELO_HA'] = df['COSTO_VUELO_HA'].apply(lambda x: x * 1000 if 0 < x < 2500 else x)
            df['COSTO_TOTAL_HA'] = df['COSTO_TOTAL_HA'].apply(lambda x: x * 1000 if 0 < x < 2500 else x)
            
            # 2. Tope Táctico: Si la tarifa de vuelo pura supera los 150.000 COP/Ha, es un error de SAP (pusieron el total). Lo topamos.
            df['COSTO_VUELO_HA'] = df['COSTO_VUELO_HA'].apply(lambda x: 75000 if x > 150000 else x)
            
            df['OPERADOR_DRON'] = df['HK'].astype(str).str.strip() + " - " + df['PISTA'].astype(str).str.strip()
            
            return df.dropna(subset=['FECHA_FILTRABLE'])
        return pd.DataFrame()
    except: return pd.DataFrame()

# =================================================================
# ⚙️ MOTOR EXCEL PROFESIONAL (CON SEMAFORIZACIÓN)
# =================================================================

def generar_excel_maestro(df_total, df_vuelo):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_total.to_excel(writer, index=False, sheet_name='Facturación Total')
        df_vuelo.to_excel(writer, index=False, sheet_name='Tarifa Vuelo Pura')
        
        header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal='center', vertical='center')
        
        font_rojo = Font(color="C00000", bold=True) 
        font_verde = Font(color="00B050", bold=True) 
        
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            
            ws.column_dimensions['A'].width = 38
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 18
            ws.column_dimensions['D'].width = 18
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 16
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                
            for row in range(2, ws.max_row + 1):
                ws[f'C{row}'].number_format = '"$"#,##0'
                ws[f'D{row}'].number_format = '"$"#,##0'
                
                celda_dif = ws[f'E{row}']
                celda_efi = ws[f'F{row}']
                
                celda_dif.number_format = '"$"#,##0'
                celda_efi.number_format = '0.0%' 
                
                if isinstance(celda_dif.value, (int, float)):
                    if celda_dif.value > 0:
                        celda_dif.font = font_verde
                        celda_efi.font = font_verde
                    elif celda_dif.value < 0:
                        celda_dif.font = font_rojo
                        celda_efi.font = font_rojo
                
    return output.getvalue()

# =================================================================
# 👑 RENDERIZADO VISUAL EN PANTALLA
# =================================================================

def ejecutar(*args, **kwargs):
    VERDE_INTENSO = '#143521'
    DORADO = '#d4af37'

    st.header("", anchor="inicio_modulo")
    
    # 🚀 CEBO DE HARDENING INDUSTRIAL: Contorno grueso perimetral de 3px Verde de Marca e Inputs Opacos
    st.markdown(f"""
    <style>
    h1 {{ color: #1a365d; font-family: Arial Black; border-bottom: 3px solid {DORADO}; }}
    div[data-testid="stDataFrame"] {{ border: 2px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }}
    
    /* Enmarcar selectores de fecha con contorno sólido de 3px color Verde Intenso */
    div[data-testid="stDateInput"] input {{
        background-color: #ffffff !important;
        border: 3px solid {VERDE_INTENSO} !important;
        border-radius: 6px !important;
    }}
    div[data-testid="stDateInput"] * {{
        color: #000000 !important;
        font-weight: bold !important;
    }}
    div[data-testid="stMainBlockContainer"] label p {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1>📊 Comparativo Financiero Detallado</h1>", unsafe_allow_html=True)

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

    df_base = df_raw[(df_raw['FECHA_FILTRABLE'] >= fecha_inicio) & (df_raw['FECHA_FILTRABLE'] <= fecha_fin)].copy()
    if df_base.empty:
        st.error(f"❌ No se encontraron registros de vuelo para el rango seleccionado.")
        return

    df_aviones = df_base[df_base['TECNOLOGIA'] == 'AVIÓN'].copy()
    df_drones = df_base[df_base['TECNOLOGIA'] == 'DRONE'].copy()

    # PREPARAR DATA: COSTO VUELO PURA
    df_vuelo_avion = df_aviones.groupby('FINCA')['COSTO_VUELO_HA'].max().reset_index().rename(columns={'COSTO_VUELO_HA': 'AVIÓN'})
    df_vuelo_dron = df_drones.groupby(['FINCA', 'OPERADOR_DRON'])['COSTO_VUELO_HA'].max().reset_index().rename(columns={'COSTO_VUELO_HA': 'DRONE'})
    m_comp_v = pd.merge(df_vuelo_dron, df_vuelo_avion, on='FINCA', how='inner')
    
    if not m_comp_v.empty:
        m_comp_v = m_comp_v[['FINCA', 'OPERADOR_DRON', 'AVIÓN', 'DRONE']]
        m_comp_v.rename(columns={'OPERADOR_DRON': 'EQUIPO DRON'}, inplace=True)
        m_comp_v['Diferencia ($)'] = m_comp_v['AVIÓN'] - m_comp_v['DRONE']
        m_comp_v['Eficiencia (%)'] = m_comp_v['Diferencia ($)'] / m_comp_v['AVIÓN']

    # PREPARAR DATA: COSTO TOTAL
    df_total_avion = df_aviones.groupby('FINCA')['COSTO_TOTAL_HA'].max().reset_index().rename(columns={'COSTO_TOTAL_HA': 'AVIÓN'})
    df_total_dron = df_drones.groupby(['FINCA', 'OPERADOR_DRON'])['COSTO_TOTAL_HA'].max().reset_index().rename(columns={'COSTO_TOTAL_HA': 'DRONE'})
    m_comp_t = pd.merge(df_total_dron, df_total_avion, on='FINCA', how='inner')

    if not m_comp_t.empty:
        m_comp_t = m_comp_t[['FINCA', 'OPERADOR_DRON', 'AVIÓN', 'DRONE']]
        m_comp_t.rename(columns={'OPERADOR_DRON': 'EQUIPO DRON'}, inplace=True)
        m_comp_t['Diferencia ($)'] = m_comp_t['AVIÓN'] - m_comp_t['DRONE']
        m_comp_t['Eficiencia (%)'] = m_comp_t['Diferencia ($)'] / m_comp_t['AVIÓN'] 

    if not m_comp_t.empty and not m_comp_v.empty:
        excel_data = generar_excel_maestro(m_comp_t, m_comp_v)
        st.download_button(
            label="📥 DESCARGAR REPORTE GERENCIAL (EXCEL A COLOR)", 
            data=excel_data, 
            file_name=f"Reporte_Eficiencia_Detallado.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    tab_vuelo, tab_total = st.tabs(["✈️ TARIFA DE VUELO PURA (Columna T)", "💰 FACTURACIÓN TOTAL OPERACIÓN"])

    def formatear_pesos(val):
        if pd.isna(val) or val == 0: return "-"
        if val < 0: return f"-$ {abs(val):,.0f}".replace(",", ".")
        return f"$ {val:,.0f}".replace(",", ".")

    def semaforo_financiero(val):
        if isinstance(val, str):
            if '-' in val and val != '-': return 'color: #e53e3e; font-weight: bold;' 
            elif val == '-': return 'color: #718096;' 
            else: return 'color: #27ae60; font-weight: bold;' 
        return ''

    # ==========================================================
    # PESTAÑA: TARIFA VUELO
    # ==========================================================
    with tab_vuelo:
        st.success("🔬 Eficiencia dividida por Equipo de Dron (Sin promedios mezclados).")
        if not m_comp_v.empty:
            df_print_v = m_comp_v.copy()
            df_print_v['AVIÓN'] = df_print_v['AVIÓN'].apply(formatear_pesos)
            df_print_v['DRONE'] = df_print_v['DRONE'].apply(formatear_pesos)
            df_print_v['Diferencia ($)'] = df_print_v['Diferencia ($ Freeman)']=df_print_v['Diferencia ($)'].apply(formatear_pesos)
            df_print_v['Eficiencia (%)'] = (df_print_v['Eficiencia (%)'] * 100).apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")

            st.dataframe(df_print_v.style.map(semaforo_financiero, subset=['Diferencia ($)', 'Eficiencia (%)']), use_container_width=True, hide_index=True)
            
            df_print_v['EJE_X'] = df_print_v['FINCA'].str[:12] + " (" + df_print_v['EQUIPO DRON'].str.split('-').str[0].str.strip() + ")"
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_print_v['EJE_X'], y=m_comp_v['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig.add_trace(go.Bar(x=df_print_v['EJE_X'], y=m_comp_v['DRONE'], name='Dron', marker_color='#d4af37'))
            fig.update_layout(
                title="Brecha Real de Tarifa Vuelo (Avión vs Dron Específico)", 
                barmode='group', 
                plot_bgcolor='rgba(0,0,0,0)', 
                xaxis=dict(tickangle=-45),
                yaxis=dict(tickformat="$,.0f", title="Costo ($ COP / ha)"), # 💥 Formato de Moneda Real sin la 'k'
                hovermode="closest" # ⚡ Restablece la interactividad responsiva
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("📌 No hay datos cruzados en el rango.")

    # ==========================================================
    # PESTAÑA: COSTO TOTAL
    # ==========================================================
    with tab_total:
        st.info("📊 Impacto macro en presupuesto desglosado por Operador (Consolidado Completo)")
        if not m_comp_t.empty:
            df_print_t = m_comp_t.copy()
            df_print_t['AVIÓN'] = df_print_t['AVIÓN'].apply(formatear_pesos)
            df_print_t['DRONE'] = df_print_t['DRONE'].apply(formatear_pesos)
            df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($)'].apply(formatear_pesos)
            df_print_t['Eficiencia (%)'] = (df_print_t['Eficiencia (%)'] * 100).apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")

            st.dataframe(df_print_t.style.map(semaforo_financiero, subset=['Diferencia ($)', 'Eficiencia (%)']), use_container_width=True, hide_index=True)
        else:
            st.warning("📌 No hay datos cruzados en el rango.")

if __name__ == "__main__":
    pass
