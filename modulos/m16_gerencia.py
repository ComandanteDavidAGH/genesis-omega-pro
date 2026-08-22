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
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =================================================================
# ⚙️ CONSTANTES Y MOTOR DE CONEXIÓN UNIFICADO (V42 VIP)
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

@st.cache_resource(show_spinner=False)
def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    if "gcp_credentials" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_credentials"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    try:
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

# =================================================================
# 🛡️ UTILIDADES DE PURIFICACIÓN Y LIMPIEZA FINANCIERA
# =================================================================
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

def procesar_fecha_pesada(val):
    if pd.isna(val) or str(val).strip() == "": return pd.NaT
    s = str(val).strip()
    if s.replace('.', '', 1).isdigit(): 
        return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y'):
        try: return pd.to_datetime(s, format=fmt)
        except: pass
    try: return pd.to_datetime(s, errors='coerce')
    except: return pd.NaT

def normalizar_a_fecha_pura(val):
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)): return res_nativo.date()
        if isinstance(res_nativo, date): return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except: return None

def es_cooperativa(finca_nombre):
    f_up = str(finca_nombre).upper().strip()
    patrones_coop = [
        'BANAFRUCOOP', 'COOMULBANANO', 'COOBAMAG', 'EMPREBANCOOP', 
        'COOBAFRIO', 'BANAORGANICO', 'COOP', 'ASOCIACION', 'ASO'
    ]
    return any(p in f_up for p in patrones_coop)

# =================================================================
# 💾 EXTRACCIÓN CACHEADA DE DATOS OPERATIVOS
# =================================================================
@st.cache_data(show_spinner=False, ttl=120)
def cargar_datos_gerenciales():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    
    try:
        boveda_act = gc.open_by_url(URL_BOVEDA_MAESTRA)
        datos_brutos = boveda_act.worksheet("TABLA 1").get_all_values()
        
        if len(datos_brutos) > 5:
            columnas_t1 = [
                "OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", 
                "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", 
                "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", 
                "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA"
            ]
            filas_limpias = [r + [""]*(len(columnas_t1) - len(r)) for r in datos_brutos[5:]]
            df = pd.DataFrame([r[:len(columnas_t1)] for r in filas_limpias], columns=columnas_t1)
            
            df['FINCA'] = df['FINCA'].astype(str).str.strip().str.upper()
            df['FECHA_FILTRABLE'] = df['FECHA'].apply(normalizar_a_fecha_pura)
            
            def clasificar_tec(row):
                texto = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))}".upper()
                if 'DRON' in texto or 'DR5' in texto: return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
            df['TIPO_ENTIDAD'] = df['FINCA'].apply(lambda x: 'COOPERATIVA' if es_cooperativa(x) else 'INDEPENDIENTE')
            
            df['COSTO_TOTAL_HA'] = df['VALOR_FACTURAR'].apply(limpiar_tarifa_excel)
            df['COSTO_VUELO_HA'] = df['COSTO_HA'].apply(limpiar_tarifa_excel)
            
            # Saneamiento de Errores de Digitación en SAP
            df['COSTO_VUELO_HA'] = df['COSTO_VUELO_HA'].apply(lambda x: x * 1000 if 0 < x < 2500 else x)
            df['COSTO_TOTAL_HA'] = df['COSTO_TOTAL_HA'].apply(lambda x: x * 1000 if 0 < x < 2500 else x)
            df['COSTO_VUELO_HA'] = df['COSTO_VUELO_HA'].apply(lambda x: 75000 if x > 150000 else x)
            
            df['OPERADOR_DRON'] = df['HK'].astype(str).str.strip() + " - " + df['PISTA'].astype(str).str.strip()
            
            return df.dropna(subset=['FECHA_FILTRABLE'])
        return pd.DataFrame()
    except Exception: return pd.DataFrame()

# =================================================================
# ⚙️ MOTOR EXCEL PROFESIONAL
# =================================================================
def generar_excel_maestro(df_total, df_vuelo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_total.to_excel(writer, index=False, sheet_name='Facturación Total')
        df_vuelo.to_excel(writer, index=False, sheet_name='Tarifa Vuelo Pura')
        
        header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        header_font = Font(color="D4AF37", bold=True)
        align_center = Alignment(horizontal='center', vertical='center')
        borde_fino = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), 
                            top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
        
        font_rojo = Font(color="C00000", bold=True) 
        font_verde = Font(color="00B050", bold=True) 
        
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            
            ws.column_dimensions['A'].width = 32
            ws.column_dimensions['B'].width = 16
            ws.column_dimensions['C'].width = 28
            ws.column_dimensions['D'].width = 18
            ws.column_dimensions['E'].width = 18
            ws.column_dimensions['F'].width = 20
            ws.column_dimensions['G'].width = 16
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                cell.border = borde_fino
                
            for row in range(2, ws.max_row + 1):
                ws[f'D{row}'].number_format = '"$"#,##0'
                ws[f'E{row}'].number_format = '"$"#,##0'
                
                celda_dif = ws[f'F{row}']
                celda_efi = ws[f'G{row}']
                
                celda_dif.number_format = '"$"#,##0'
                celda_efi.number_format = '0.0%' 
                
                for col_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                    ws[f'{col_letter}{row}'].border = borde_fino
                
                if isinstance(celda_dif.value, (int, float)):
                    if celda_dif.value > 0:
                        celda_dif.font = font_verde
                        celda_efi.font = font_verde
                    elif celda_dif.value < 0:
                        celda_dif.font = font_rojo
                        celda_efi.font = font_rojo
                
    return output.getvalue()

def construir_grafico_comparativo(df_datos, titulo_grafico):
    if df_datos.empty:
        return None
        
    df_plot = df_datos.copy()
    df_plot['X_UNIQUE'] = df_plot['FINCA'].apply(lambda x: str(x)[:14] + '...' if len(str(x)) > 14 else str(x)) + " [" + df_plot.index.astype(str) + "]"
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_plot['X_UNIQUE'], 
        y=df_plot['AVIÓN'], 
        name='Avión', 
        marker_color='#1a365d',
        customdata=df_plot['FINCA'],
        hovertemplate='<b>Finca:</b> %{customdata}<br><b>Avión:</b> $%{y:,.0f}<extra></extra>'
    ))
    
    fig.add_trace(go.Bar(
        x=df_plot['X_UNIQUE'], 
        y=df_plot['DRONE'], 
        name='Dron', 
        marker_color='#d4af37',
        customdata=df_plot['FINCA'],
        hovertemplate='<b>Finca:</b> %{customdata}<br><b>Dron:</b> $%{y:,.0f}<extra></extra>'
    ))
    
    vista_inicial = min(12.5, len(df_plot) - 0.5) 
    
    fig.update_layout(
        title=f"<b>{titulo_grafico}</b>", 
        barmode='group', 
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)',
        height=420,
        xaxis=dict(
            tickangle=-90,
            tickfont=dict(size=10),
            range=[-0.5, vista_inicial],
            rangeslider=dict(visible=True, thickness=0.06, bgcolor="#e2e8f0"),
            type='category'
        ),
        yaxis=dict(
            tickformat="$,.0f", 
            title="Costo ($ COP / ha)", 
            showgrid=True, 
            gridcolor='rgba(200,200,200,0.2)'
        ),
        hovermode="closest",
        margin=dict(b=10, t=40, l=10, r=10)
    )
    return fig

# =================================================================
# 👑 RENDERIZADO VISUAL EN PANTALLA
# =================================================================
def ejecutar(*args, **kwargs):
    VERDE_INTENSO = '#143521'
    DORADO = '#d4af37'

    st.markdown(f"""
    <style>
    .titulo-gerencial {{ color: #0d1b2a; border-bottom: 3px solid {DORADO}; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }}
    
    [data-testid="column"] {{ display: flex !important; flex-direction: column !important; justify-content: flex-start !important; align-items: stretch !important; }}
    div[data-testid="stDataFrame"] {{ border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.08) !important; }}
    
    div[data-testid="stDateInput"] input {{ background-color: #ffffff !important; border: 2px solid {VERDE_INTENSO} !important; border-radius: 6px !important; }}
    div[data-testid="stDateInput"] * {{ color: #000000 !important; font-weight: bold !important; }}
    div[data-testid="stMainBlockContainer"] label p {{ color: #0d1b2a !important; font-weight: 800 !important; text-transform: uppercase !important; }}
    
    div[data-testid="stTabs"] button[role="tab"] {{ font-family: 'Arial Black', sans-serif; font-size: 14px; color: #0d1b2a; }}
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{ border-bottom-color: {DORADO}; background-color: rgba(212, 175, 55, 0.1); }}
    
    [data-testid="stPlotlyChart"] {{ transition: transform 0.3s ease, box-shadow 0.3s ease !important; border-radius: 8px; }}
    [data-testid="stPlotlyChart"]:hover {{ transform: translateY(-4px) scale(1.015) !important; box-shadow: 0 12px 25px rgba(212, 175, 55, 0.25) !important; z-index: 10; }}
    </style>
    """, unsafe_allow_html=True)

    def tarjeta_kpi(titulo, valor, delta_texto="", color_delta="#28a745"):
        delta_html = f"<span style='font-size: 14px; color: {color_delta}; margin-left: 8px; vertical-align: middle; padding: 2px 6px; border-radius: 4px; background-color: rgba(255,255,255,0.1);'>{delta_texto}</span>" if delta_texto else ""
        return f"""
        <div style='background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid {DORADO}; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 20px; height: 100%; min-height: 85px; display: flex; flex-direction: column; justify-content: center;'>
            <p style='font-size: 11px; font-weight: bold; color: {DORADO}; text-transform: uppercase; margin:0 0 5px 0; letter-spacing: 1px;'>{titulo}</p>
            <p style='font-size: 22px; font-family: "Arial Black", sans-serif; margin: 0; color: white; display: flex; align-items: center;'>{valor} {delta_html}</p>
        </div>
        """

    c_tit, c_sync = st.columns([3.5, 1.5])
    with c_tit:
        st.markdown("<h1 class='titulo-gerencial'>⚖️ Módulo 16: Comparativo Gerencial (Dron vs Avión)</h1>", unsafe_allow_html=True)
        st.write("Análisis táctico de costos desglosado por Cooperativas vs Fincas Independientes.")
    with c_sync:
        st.write("")
        if st.button("🔄 Sincronizar Base Datos", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.rerun()

    with st.container(border=True):
        st.markdown("#### 📅 Parámetros de Análisis")
        c_f1, c_f2 = st.columns(2)
        fecha_inicio = c_f1.date_input("Desde:", value=date(2026, 1, 1))
        fecha_fin = c_f2.date_input("Hasta:", value=date(2026, 12, 31))

    df_raw = cargar_datos_gerenciales()
    if df_raw.empty:
        st.warning("⚠️ No se detectan registros en la base maestra.")
        return

    df_base = df_raw[(df_raw['FECHA_FILTRABLE'] >= fecha_inicio) & (df_raw['FECHA_FILTRABLE'] <= fecha_fin)].copy()
    if df_base.empty:
        st.error("❌ No se encontraron registros de vuelo para el rango seleccionado.")
        return

    def purificar_tarifa(val, tope_max, valor_reemplazo):
        if pd.isna(val): return 0.0
        if 0 < val < 2500: val = val * 1000
        if val > tope_max: return valor_reemplazo
        return val

    df_base['COSTO_VUELO_HA'] = df_base['COSTO_VUELO_HA'].apply(lambda x: purificar_tarifa(x, 150000, 75000))
    df_base['COSTO_TOTAL_HA'] = df_base['COSTO_TOTAL_HA'].apply(lambda x: purificar_tarifa(x, 400000, 200000))

    df_aviones = df_base[df_base['TECNOLOGIA'] == 'AVIÓN'].copy()
    df_drones = df_base[df_base['TECNOLOGIA'] == 'DRONE'].copy()

    # PREPARAR DATA: COSTO VUELO PURA
    df_vuelo_avion = df_aviones.groupby(['FINCA', 'TIPO_ENTIDAD'])['COSTO_VUELO_HA'].max().reset_index().rename(columns={'COSTO_VUELO_HA': 'AVIÓN'})
    df_vuelo_dron = df_drones.groupby(['FINCA', 'TIPO_ENTIDAD', 'OPERADOR_DRON'])['COSTO_VUELO_HA'].max().reset_index().rename(columns={'COSTO_VUELO_HA': 'DRONE'})
    m_comp_v = pd.merge(df_vuelo_dron, df_vuelo_avion, on=['FINCA', 'TIPO_ENTIDAD'], how='inner')
    
    if not m_comp_v.empty:
        m_comp_v = m_comp_v[['FINCA', 'TIPO_ENTIDAD', 'OPERADOR_DRON', 'AVIÓN', 'DRONE']]
        m_comp_v.rename(columns={'OPERADOR_DRON': 'EQUIPO DRON'}, inplace=True)
        m_comp_v['Diferencia ($)'] = m_comp_v['AVIÓN'] - m_comp_v['DRONE']
        m_comp_v['Eficiencia (%)'] = m_comp_v['Diferencia ($)'] / m_comp_v['AVIÓN']

    # PREPARAR DATA: COSTO TOTAL
    df_total_avion = df_aviones.groupby(['FINCA', 'TIPO_ENTIDAD'])['COSTO_TOTAL_HA'].max().reset_index().rename(columns={'COSTO_TOTAL_HA': 'AVIÓN'})
    df_total_dron = df_drones.groupby(['FINCA', 'TIPO_ENTIDAD', 'OPERADOR_DRON'])['COSTO_TOTAL_HA'].max().reset_index().rename(columns={'COSTO_TOTAL_HA': 'DRONE'})
    m_comp_t = pd.merge(df_total_dron, df_total_avion, on=['FINCA', 'TIPO_ENTIDAD'], how='inner')

    if not m_comp_t.empty:
        m_comp_t = m_comp_t[['FINCA', 'TIPO_ENTIDAD', 'OPERADOR_DRON', 'AVIÓN', 'DRONE']]
        m_comp_t.rename(columns={'OPERADOR_DRON': 'EQUIPO DRON'}, inplace=True)
        m_comp_t['Diferencia ($)'] = m_comp_t['AVIÓN'] - m_comp_t['DRONE']
        m_comp_t['Eficiencia (%)'] = m_comp_t['Diferencia ($)'] / m_comp_t['AVIÓN'] 

    # ==========================================================
    # 💎 TARJETAS KPI DE IMPACTO DIRECTO
    # ==========================================================
    st.markdown("---")
    if not m_comp_v.empty:
        ahorro_prom_vuelo = m_comp_v['Diferencia ($)'].mean()
        eficiencia_prom_vuelo = m_comp_v['Eficiencia (%)'].mean() * 100
        fincas_coop = m_comp_v[m_comp_v['TIPO_ENTIDAD'] == 'COOPERATIVA']['FINCA'].nunique()
        fincas_indep = m_comp_v[m_comp_v['TIPO_ENTIDAD'] == 'INDEPENDIENTE']['FINCA'].nunique()

        k1, k2, k3 = st.columns(3)
        with k1: st.markdown(tarjeta_kpi("Cobertura Mapeada", f"{fincas_coop} Coop / {fincas_indep} Indep", "Fincas Cruzadas", "#d4af37"), unsafe_allow_html=True)
        with k2: st.markdown(tarjeta_kpi("Brecha Promedio Vuelo", f"$ {ahorro_prom_vuelo:,.0f} /ha".replace(",", "."), "Ahorro Dron vs Avión", "#28a745" if ahorro_prom_vuelo >= 0 else "#dc3545"), unsafe_allow_html=True)
        with k3: st.markdown(tarjeta_kpi("Eficiencia Financiera", f"{eficiencia_prom_vuelo:.1f}%", "vs Tarifa Avión", "#28a745" if eficiencia_prom_vuelo >= 0 else "#dc3545"), unsafe_allow_html=True)

    tab_vuelo, tab_total = st.tabs(["✈️ Tarifa Vuelo Pura (Operativo)", "💰 Facturación Total Operación"])

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
    # PESTAÑA 1: TARIFA VUELO PURA (SEGMENTADA EN 2 GRÁFICOS)
    # ==========================================================
    with tab_vuelo:
        st.success("🔬 Análisis de Tarifas Vuelo Pura: Cooperativas vs Fincas Independientes.")
        
        if not m_comp_v.empty:
            m_comp_v_coop = m_comp_v[m_comp_v['TIPO_ENTIDAD'] == 'COOPERATIVA'].copy()
            m_comp_v_indep = m_comp_v[m_comp_v['TIPO_ENTIDAD'] == 'INDEPENDIENTE'].copy()
            
            # --- 🏢 GRÁFICO 1: COOPERATIVAS ---
            st.markdown("### 🏢 1. Tarifas en Cooperativas y Gremios")
            if not m_comp_v_coop.empty:
                fig_coop = construir_grafico_comparativo(m_comp_v_coop, "Cooperativas: Avión vs Dron")
                if fig_coop: st.plotly_chart(fig_coop, use_container_width=True)
            else:
                st.info("No se registraron fincas asociadas a Cooperativas en el rango.")

            st.markdown("<br>", unsafe_allow_html=True)

            # --- 🚜 GRÁFICO 2: FINCAS INDEPENDIENTES ---
            st.markdown("### 🚜 2. Tarifas en Fincas Normales / Independientes")
            if not m_comp_v_indep.empty:
                fig_indep = construir_grafico_comparativo(m_comp_v_indep, "Fincas Independientes: Avión vs Dron")
                if fig_indep: st.plotly_chart(fig_indep, use_container_width=True)
            else:
                st.info("No se registraron Fincas Independientes en el rango.")

            # TABLA GENERAL DETALLADA
            st.markdown("---")
            st.markdown("#### 📋 Matriz Detallada de Comparación")
            df_print_v = m_comp_v.copy()
            df_print_v['AVIÓN'] = df_print_v['AVIÓN'].apply(formatear_pesos)
            df_print_v['DRONE'] = df_print_v['DRONE'].apply(formatear_pesos)
            df_print_v['Diferencia ($)'] = df_print_v['Diferencia ($)'].apply(formatear_pesos)
            df_print_v['Eficiencia (%)'] = (df_print_v['Eficiencia (%)'] * 100).apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")

            st.dataframe(
                df_print_v.style.map(semaforo_financiero, subset=['Diferencia ($)', 'Eficiencia (%)']), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.warning("📌 No hay datos cruzados en el rango seleccionado.")

    # ==========================================================
    # PESTAÑA 2: COSTO TOTAL FACTURADO (SEGMENTADA EN 2 GRÁFICOS)
    # ==========================================================
    with tab_total:
        st.info("📊 Impacto Macro en Facturación Total: Cooperativas vs Fincas Independientes.")
        
        if not m_comp_t.empty:
            m_comp_t_coop = m_comp_t[m_comp_t['TIPO_ENTIDAD'] == 'COOPERATIVA'].copy()
            m_comp_t_indep = m_comp_t[m_comp_t['TIPO_ENTIDAD'] == 'INDEPENDIENTE'].copy()
            
            # --- 🏢 GRÁFICO 1: COOPERATIVAS (TOTAL) ---
            st.markdown("### 🏢 1. Facturación Total en Cooperativas")
            if not m_comp_t_coop.empty:
                fig_t_coop = construir_grafico_comparativo(m_comp_t_coop, "Facturación Total: Cooperativas")
                if fig_t_coop: st.plotly_chart(fig_t_coop, use_container_width=True)
            else:
                st.info("No se registraron Cooperativas en el rango.")

            st.markdown("<br>", unsafe_allow_html=True)

            # --- 🚜 GRÁFICO 2: FINCAS INDEPENDIENTES (TOTAL) ---
            st.markdown("### 🚜 2. Facturación Total en Fincas Independientes")
            if not m_comp_t_indep.empty:
                fig_t_indep = construir_grafico_comparativo(m_comp_t_indep, "Facturación Total: Fincas Independientes")
                if fig_t_indep: st.plotly_chart(fig_t_indep, use_container_width=True)
            else:
                st.info("No se registraron Fincas Independientes en el rango.")

            # TABLA GENERAL DETALLADA TOTAL
            st.markdown("---")
            st.markdown("#### 📋 Matriz Detallada Facturación Total")
            df_print_t = m_comp_t.copy()
            df_print_t['AVIÓN'] = df_print_t['AVIÓN'].apply(formatear_pesos)
            df_print_t['DRONE'] = df_print_t['DRONE'].apply(formatear_pesos)
            df_print_t['Diferencia ($)'] = df_print_t['Diferencia ($)'].apply(formatear_pesos)
            df_print_t['Eficiencia (%)'] = (df_print_t['Eficiencia (%)'] * 100).apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")

            st.dataframe(
                df_print_t.style.map(semaforo_financiero, subset=['Diferencia ($)', 'Eficiencia (%)']), 
                use_container_width=True, 
                hide_index=True
            )
        else:
            st.warning("📌 No hay datos cruzados en el rango seleccionado.")

    # Botón de Descarga Excel VIP
    if not m_comp_t.empty and not m_comp_v.empty:
        st.markdown("---")
        excel_data = generar_excel_maestro(m_comp_t, m_comp_v)
        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL EN EXCEL (2 HOJAS CON TIPO DE ENTIDAD Y SEMÁFORO)", 
            data=excel_data, 
            file_name=f"Reporte_Eficiencia_Avion_vs_Dron.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            type="primary",
            use_container_width=True
        )

if __name__ == "__main__":
    pass
