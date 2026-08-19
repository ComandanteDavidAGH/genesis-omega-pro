import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 MOTORES DE CONEXIÓN Y FORMATO
# =================================================================

def formato_latino(numero, decimales=0):
    if pd.isna(numero) or numero is None: return "0"
    try:
        num = float(numero)
        if num == 0: return "0"
        if decimales == 0: texto_us = f"{num:,.0f}"
        else: texto_us = f"{num:,.{decimales}f}"
        return texto_us.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0"

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        elif "gcp_credentials" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception: return None

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

def limpiar_dinero(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip()
        if not v: return 0.0
        v = v.replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1:
            partes = v.rsplit('.', 1)
            v = partes[0].replace('.', '') + '.' + partes[1]
        num = float(v) if v else 0.0
        if 5 < num < 2000: num = num * 1000
        return num
    except: return 0.0

def procesar_fecha_estricta(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() in ["none", "nan", "nat", "<na>"]: return pd.NaT
    s = str(val).strip().lower()
    if s.replace('.', '', 1).isdigit(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
    meses_es = {'enero':1, 'febrero':2, 'marzo':3, 'abril':4, 'mayo':5, 'junio':6, 'julio':7, 'agosto':8, 'septiembre':9, 'octubre':10, 'noviembre':11, 'diciembre':12}
    mes_encontrado = next((meses_es[m] for m in meses_es if m in s), None)
    if mes_encontrado:
        numeros = re.findall(r'\d+', s)
        if len(numeros) >= 2:
            n1, n2 = int(numeros[0]), int(numeros[1])
            anio = n1 if n1 > 1000 else (n2 if n2 > 1000 else (2000 + n2 if n2 < 100 else n2))
            dia = n2 if n1 > 1000 else (n1 if n2 > 1000 else n1)
            try: return pd.Timestamp(year=anio, month=mes_encontrado, day=dia)
            except: pass
    s = s.replace(',', '').replace(' de ', '/').replace('-', '/').strip()
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y'):
        try: return pd.to_datetime(s, format=fmt)
        except: pass
    try: 
        res = pd.to_datetime(s, dayfirst=True)
        return pd.NaT if pd.isna(res) else res
    except: return pd.NaT 

# 💥 EXTRACCIÓN MAESTRA DEL HISTÓRICO Y TABLA 1 VIVA
@st.cache_data(show_spinner=False, ttl=600)
def cargar_fuentes_maestras_duelo():
    gc = inicializar_cliente_gspread()
    if not gc: return pd.DataFrame()
    
    # --- TABLA 1 (VIVA) ---
    try:
        boveda_act = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        datos_brutos_act = boveda_act.worksheet("TABLA 1").get_all_values()
    except:
        datos_brutos_act = []
    
    df_vivos = pd.DataFrame()
    if len(datos_brutos_act) > 5:
        columnas_t1 = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA", "INC_2026", "LIMITE", "ALERTA", "VAR_PCT", "COSTO_TOTAL", "PAGO_AVION"]
        filas_limpias = [r + [""]*(len(columnas_t1) - len(r)) for r in datos_brutos_act[5:]]
        df_vivos = pd.DataFrame([r[:len(columnas_t1)] for r in filas_limpias], columns=columnas_t1)
        df_vivos.rename(columns={'AREA_FUMIG': 'AREA_MAESTRA', 'COSTO_HA': 'COSTO_HA_BASE', 'VALOR_FACTURAR': 'COSTO_OS_TOTAL', 'FINCA': 'FINCA_MAESTRA', 'FECHA': 'FECHA_MAESTRA', 'PISTA': 'PISTA_MAESTRA', 'OS': 'OS_MAESTRA'}, inplace=True)
        df_vivos['ORIGEN'] = 'ACTUAL'

    # --- HISTÓRICO ---
    datos_brutos_hist = []
    try:
        boveda_hist = gc.open_by_url("https://docs.google.com/spreadsheets/d/16OZdiWwW7nLHyZBEnhiKlDTDttR7Tjhn37O9zm6wJOk/edit")
        datos_brutos_hist = boveda_hist.worksheet("Datos").get_all_values()
    except: pass
    
    df_historico = pd.DataFrame()
    if len(datos_brutos_hist) > 0:
        df_historico = pd.DataFrame(datos_brutos_hist[1:], columns=datos_brutos_hist[0])
        # Limpieza de encabezados históricos
        df_historico.columns = [str(c).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').strip() for c in df_historico.columns]
        renombres = {}
        for col in df_historico.columns:
            if 'FUMIG' in col and 'AREA' in col: renombres[col] = 'AREA_MAESTRA'
            elif 'FACTURAR' in col and 'VALOR' in col: renombres[col] = 'COSTO_OS_TOTAL'
            elif 'COSTO' in col and 'HA' in col: renombres[col] = 'COSTO_HA_BASE'
            elif col in ['FINCA', 'PROPIEDAD']: renombres[col] = 'FINCA_MAESTRA'
            elif col == 'FECHA': renombres[col] = 'FECHA_MAESTRA'
            elif col == 'PISTA': renombres[col] = 'PISTA_MAESTRA'
            elif "ORDEN" in col or "OS" == col: renombres[col] = 'OS_MAESTRA'
        df_historico.rename(columns=renombres, inplace=True)
        df_historico['ORIGEN'] = 'HISTORICO'

    super_base = pd.concat([df_historico, df_vivos], ignore_index=True)
    
    if not super_base.empty and 'FINCA_MAESTRA' in super_base.columns:
        super_base['FINCA_MAESTRA'] = super_base['FINCA_MAESTRA'].astype(str).str.strip().str.upper()
        super_base['PISTA_MAESTRA'] = super_base['PISTA_MAESTRA'].astype(str).str.strip().str.upper()
        super_base['FECHA_DT'] = super_base['FECHA_MAESTRA'].apply(procesar_fecha_estricta)
        super_base = super_base.dropna(subset=['FECHA_DT'])
        
        super_base['AREA_NUM'] = super_base['AREA_MAESTRA'].apply(limpiar_area)
        super_base['COSTO_HA_NUM'] = super_base['COSTO_HA_BASE'].apply(limpiar_dinero)
        super_base['COSTO_TOTAL_OS'] = super_base['COSTO_OS_TOTAL'].apply(limpiar_dinero)
        
        # Filtro de purga para datos basura
        super_base = super_base[super_base['AREA_NUM'] > 0]
        super_base = super_base[super_base['FINCA_MAESTRA'] != ""]
        
        return super_base
    else:
        return pd.DataFrame()

# =================================================================
# 👑 INTERFAZ PRINCIPAL (EL COLISEO)
# =================================================================

def ejecutar():
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    .kpi-vs { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); color: white; padding: 20px; border-radius: 12px; border-left: 6px solid #d4af37; box-shadow: 0 8px 16px rgba(0,0,0,0.3); text-align: center; margin-bottom: 15px; }
    .kpi-vs-title { font-size: 13px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin-bottom: 5px; letter-spacing: 1px; }
    .kpi-vs-value { font-size: 32px; font-weight: 900; margin: 0; color: #ffffff; font-family: 'Arial Black', sans-serif; }
    .kpi-vs-sub { font-size: 14px; font-weight: bold; color: #a0aec0; margin-top: 5px; }
    .victoria { border: 3px solid #28a745 !important; box-shadow: 0 0 15px rgba(40, 167, 69, 0.5) !important; }
    .derrota { border: 3px solid #dc3545 !important; }
    div[data-testid="stSelectbox"] > div:last-child { border: 2px solid #0d1b2a !important; border-radius: 8px !important; background-color: #ffffff !important; font-weight:bold !important;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-mod'>⚔️ 20. Duelo de Titanes (Finca vs Finca)</h1>", unsafe_allow_html=True)
    st.write("Bienvenido al Coliseo Operativo. Compara el rendimiento financiero y operativo de dos fincas cara a cara para descubrir los verdaderos márgenes.")

    with st.spinner("Desplegando el radar sobre la Bóveda Maestra de Vuelos..."):
        df_base = cargar_fuentes_maestras_duelo()

    if df_base.empty:
        st.error("🚨 No se encontró información en la base de datos o hubo un error de conexión.")
        return

    # --- FILTROS DEL COLISEO ---
    st.markdown("### 🎯 Configuración del Duelo")
    
    lista_fincas = sorted(df_base['FINCA_MAESTRA'].unique().tolist())
    
    col_filt1, col_filt2, col_filt3 = st.columns([1, 1, 1])
    
    with col_filt1:
        finca_A = st.selectbox("🔴 RETADOR A (Finca 1)", lista_fincas, index=0)
    with col_filt2:
        finca_B = st.selectbox("🔵 RETADOR B (Finca 2)", lista_fincas, index=1 if len(lista_fincas) > 1 else 0)
    with col_filt3:
        fechas_unicas = df_base['FECHA_DT'].dt.year.unique()
        opciones_tiempo = ["TODO EL HISTÓRICO"] + sorted([str(x) for x in fechas_unicas], reverse=True)
        tiempo_sel = st.selectbox("📅 Ventana de Tiempo", opciones_tiempo)

    # --- FILTRADO DE DATOS ---
    df_A = df_base[df_base['FINCA_MAESTRA'] == finca_A]
    df_B = df_base[df_base['FINCA_MAESTRA'] == finca_B]
    
    if tiempo_sel != "TODO EL HISTÓRICO":
        df_A = df_A[df_A['FECHA_DT'].dt.year == int(tiempo_sel)]
        df_B = df_B[df_B['FECHA_DT'].dt.year == int(tiempo_sel)]

    st.markdown("<hr style='border: 1px solid #d4af37;'>", unsafe_allow_html=True)

    if df_A.empty and df_B.empty:
        st.warning("⚠️ Ninguna de las fincas tiene vuelos registrados en el periodo seleccionado.")
        return

    # --- CÁLCULO DE KPIs ---
    def calcular_metricas(df_finca):
        if df_finca.empty: return 0, 0, 0, 0, "N/A"
        
        total_ha = df_finca['AREA_NUM'].sum()
        total_inversion = df_finca['COSTO_TOTAL_OS'].sum()
        costo_promedio_ha = df_finca['COSTO_HA_NUM'].mean()
        total_vuelos = df_finca['OS_MAESTRA'].nunique()
        
        # Pista Favorita (Donde más se voló)
        try: pista_fav = df_finca.groupby('PISTA_MAESTRA')['AREA_NUM'].sum().idxmax()
        except: pista_fav = "N/A"
        
        return total_ha, total_inversion, costo_promedio_ha, total_vuelos, pista_fav

    ha_A, inv_A, costo_ha_A, vuelos_A, pista_A = calcular_metricas(df_A)
    ha_B, inv_B, costo_ha_B, vuelos_B, pista_B = calcular_metricas(df_B)

    # Lógica de Victoria (El más barato gana)
    clase_win_A = "victoria" if (costo_ha_A < costo_ha_B and costo_ha_A > 0) or costo_ha_B == 0 else ""
    clase_win_B = "victoria" if (costo_ha_B < costo_ha_A and costo_ha_B > 0) or costo_ha_A == 0 else ""

    # --- RENDERIZADO DEL CUADRILÁTERO ---
    col_A, col_vs, col_B = st.columns([4, 1, 4])

    with col_A:
        st.markdown(f"<h3 style='text-align:center; color:#dc3545;'>🔴 {finca_A}</h3>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-vs {clase_win_A}'><p class='kpi-vs-title'>Costo Promedio x Hectárea</p><p class='kpi-vs-value'>$ {formato_latino(costo_ha_A, 0)}</p></div>", unsafe_allow_html=True)
        
        m_a1, m_a2 = st.columns(2)
        m_a1.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Total Facturado (OS)</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(inv_A, 0)}</p></div>", unsafe_allow_html=True)
        m_a2.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Área Fumigada</p><p class='kpi-vs-value' style='font-size:20px;'>{formato_latino(ha_A, 1)} Ha</p></div>", unsafe_allow_html=True)
        
        st.info(f"**🛫 Pista Dominante:** {pista_A} ({vuelos_A} Vuelos Registrados)")

    with col_vs:
        st.markdown("<br><br><h1 style='text-align:center; color:#d4af37; font-size: 50px; font-family:Arial Black;'>VS</h1>", unsafe_allow_html=True)

    with col_B:
        st.markdown(f"<h3 style='text-align:center; color:#2F75B5;'>🔵 {finca_B}</h3>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-vs {clase_win_B}'><p class='kpi-vs-title'>Costo Promedio x Hectárea</p><p class='kpi-vs-value'>$ {formato_latino(costo_ha_B, 0)}</p></div>", unsafe_allow_html=True)
        
        m_b1, m_b2 = st.columns(2)
        m_b1.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Total Facturado (OS)</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(inv_B, 0)}</p></div>", unsafe_allow_html=True)
        m_b2.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Área Fumigada</p><p class='kpi-vs-value' style='font-size:20px;'>{formato_latino(ha_B, 1)} Ha</p></div>", unsafe_allow_html=True)

        st.info(f"**🛫 Pista Dominante:** {pista_B} ({vuelos_B} Vuelos Registrados)")

    st.markdown("<hr>", unsafe_allow_html=True)

    # --- ANÁLISIS DE PISTAS (RENDIMIENTO LOGÍSTICO) ---
    st.markdown("### 🛫 Rendimiento Logístico por Pista")
    st.caption("Descubre si el costo varía dependiendo de la pista de donde despega el avión para cada finca.")

    df_ambos = pd.concat([df_A, df_B])
    
    if not df_ambos.empty:
        # Agrupar por Finca y Pista
        df_pistas = df_ambos.groupby(['FINCA_MAESTRA', 'PISTA_MAESTRA']).agg(
            COSTO_PROMEDIO=('COSTO_HA_NUM', 'mean'),
            VUELOS=('OS_MAESTRA', 'nunique')
        ).reset_index()

        df_pistas = df_pistas[df_pistas['PISTA_MAESTRA'] != ""]

        fig_pistas = px.bar(
            df_pistas, 
            x='PISTA_MAESTRA', 
            y='COSTO_PROMEDIO', 
            color='FINCA_MAESTRA', 
            barmode='group',
            text='COSTO_PROMEDIO',
            color_discrete_map={finca_A: '#dc3545', finca_B: '#2F75B5'},
            title="Comparativo de Costo / Hectárea según Base de Despegue"
        )
        fig_pistas.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont=dict(family="Arial Black"))
        fig_pistas.update_layout(
            yaxis_title="Costo Promedio ($ / Ha)", 
            xaxis_title="Pista de Operación",
            plot_bgcolor='rgba(0,0,0,0)', 
            paper_bgcolor='#ffffff',
            legend_title_text='Finca'
        )
        st.plotly_chart(fig_pistas, use_container_width=True)

    # --- CARRERA EN EL TIEMPO (TENDENCIA) ---
    st.markdown("### 📈 Carrera Operativa: Evolución del Costo")
    st.caption("Observa cómo se ha comportado el precio a lo largo de los meses para ambas fincas.")

    if not df_ambos.empty:
        df_ambos['MES_AÑO'] = df_ambos['FECHA_DT'].dt.to_period('M').astype(str)
        df_tendencia = df_ambos.groupby(['MES_AÑO', 'FINCA_MAESTRA'])['COSTO_HA_NUM'].mean().reset_index()
        df_tendencia = df_tendencia.sort_values(by='MES_AÑO')

        fig_line = px.line(
            df_tendencia, 
            x='MES_AÑO', 
            y='COSTO_HA_NUM', 
            color='FINCA_MAESTRA',
            markers=True,
            color_discrete_map={finca_A: '#dc3545', finca_B: '#2F75B5'},
            title="Línea de Tiempo: Costo por Hectárea"
        )
        fig_line.update_traces(line=dict(width=4), marker=dict(size=10))
        fig_line.update_layout(
            yaxis_title="Costo ($ / Ha)", 
            xaxis_title="Mes de Operación",
            plot_bgcolor='rgba(0,0,0,0)', 
            paper_bgcolor='#ffffff',
            hovermode="x unified"
        )
        st.plotly_chart(fig_line, use_container_width=True)

if __name__ == "__main__":
    pass
