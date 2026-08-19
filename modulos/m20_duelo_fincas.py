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
        
        # 🛡️ ESCUDO ANTI-DUPLICADOS T1
        df_vivos = df_vivos.loc[:, ~df_vivos.columns.duplicated()]

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
        
        # 🛡️ ESCUDO ANTI-DUPLICADOS HISTÓRICO (Elimina columnas sin nombre o repetidas)
        df_historico = df_historico.loc[:, ~df_historico.columns.duplicated()]
        df_historico = df_historico.loc[:, [c for c in df_historico.columns if c != ""]]

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
# 👑 INTERFAZ PRINCIPAL (EL COLISEO LOGÍSTICO)
# =================================================================

def ejecutar():
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    .kpi-vs { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); color: white; padding: 20px; border-radius: 12px; border-left: 6px solid #d4af37; box-shadow: 0 8px 16px rgba(0,0,0,0.3); text-align: center; margin-bottom: 15px; }
    .kpi-vs-title { font-size: 13px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin-bottom: 5px; letter-spacing: 1px; }
    .kpi-vs-value { font-size: 32px; font-weight: 900; margin: 0; color: #ffffff; font-family: 'Arial Black', sans-serif; }
    .victoria { border: 3px solid #28a745 !important; box-shadow: 0 0 15px rgba(40, 167, 69, 0.5) !important; }
    div[data-testid="stSelectbox"] > div:last-child { border: 2px solid #0d1b2a !important; border-radius: 8px !important; background-color: #ffffff !important; font-weight:bold !important;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-mod'>⚔️ 20. Duelo Logístico (Pista vs Pista)</h1>", unsafe_allow_html=True)
    st.write("Analiza el rendimiento de una misma finca operando desde dos bases distintas. Descubre qué pista ofrece el mejor costo por Orden de Servicio y Hectárea para tu cliente.")

    with st.spinner("Desplegando el radar sobre la Bóveda Maestra de Vuelos..."):
        df_base = cargar_fuentes_maestras_duelo()

    if df_base.empty:
        st.error("🚨 No se encontró información en la base de datos o hubo un error de conexión.")
        return

    # --- SELECCIÓN DE LA FINCA ---
    st.markdown("### 🎯 Configuración del Duelo")
    
    lista_fincas = sorted(df_base['FINCA_MAESTRA'].unique().tolist())
    finca_sel = st.selectbox("🏡 SELECCIONE LA FINCA A ANALIZAR", lista_fincas)

    # Filtramos la base solo para esa finca
    df_finca_global = df_base[df_base['FINCA_MAESTRA'] == finca_sel]
    lista_pistas = sorted(df_finca_global['PISTA_MAESTRA'].unique().tolist())

    if not lista_pistas:
        st.warning("⚠️ No hay pistas registradas para esta finca.")
        return

    col_filt1, col_filt2, col_filt3 = st.columns([1, 1, 1])
    
    with col_filt1:
        pista_A = st.selectbox("🔴 PISTA RETADORA A", lista_pistas, index=0)
    with col_filt2:
        pista_B = st.selectbox("🔵 PISTA RETADORA B", lista_pistas, index=1 if len(lista_pistas) > 1 else 0)
    with col_filt3:
        fechas_unicas = df_finca_global['FECHA_DT'].dt.year.unique()
        opciones_tiempo = ["TODO EL HISTÓRICO"] + sorted([str(x) for x in fechas_unicas], reverse=True)
        tiempo_sel = st.selectbox("📅 Ventana de Tiempo", opciones_tiempo)

    # --- FILTRADO DE DATOS (Por Tiempo y Pista) ---
    df_finca = df_finca_global.copy()
    if tiempo_sel != "TODO EL HISTÓRICO":
        df_finca = df_finca[df_finca['FECHA_DT'].dt.year == int(tiempo_sel)]

    df_A = df_finca[df_finca['PISTA_MAESTRA'] == pista_A]
    df_B = df_finca[df_finca['PISTA_MAESTRA'] == pista_B]

    st.markdown("<hr style='border: 1px solid #d4af37;'>", unsafe_allow_html=True)

    if df_A.empty and df_B.empty:
        st.warning("⚠️ La finca no operó desde ninguna de estas pistas en el periodo seleccionado.")
        return

    # --- CÁLCULO DE KPIs ---
    def calcular_metricas(df_pista):
        if df_pista.empty: return 0, 0, 0, 0, 0
        
        total_ha = df_pista['AREA_NUM'].sum()
        total_inversion = df_pista['COSTO_TOTAL_OS'].sum()
        costo_promedio_ha = df_pista['COSTO_HA_NUM'].mean()
        total_vuelos = df_pista['OS_MAESTRA'].nunique()
        # Nuevo cálculo solicitado: Costo Promedio por Orden de Servicio
        costo_promedio_os = total_inversion / total_vuelos if total_vuelos > 0 else 0
        
        return total_ha, total_inversion, costo_promedio_ha, total_vuelos, costo_promedio_os

    ha_A, inv_A, costo_ha_A, vuelos_A, costo_os_A = calcular_metricas(df_A)
    ha_B, inv_B, costo_ha_B, vuelos_B, costo_os_B = calcular_metricas(df_B)

    # Lógica de Victoria (El Costo Promedio por Ha más barato gana)
    clase_win_A = "victoria" if (costo_ha_A < costo_ha_B and costo_ha_A > 0) or costo_ha_B == 0 else ""
    clase_win_B = "victoria" if (costo_ha_B < costo_ha_A and costo_ha_B > 0) or costo_ha_A == 0 else ""

    # --- RENDERIZADO DEL CUADRILÁTERO ---
    col_A, col_vs, col_B = st.columns([4, 1, 4])

    with col_A:
        st.markdown(f"<h3 style='text-align:center; color:#dc3545;'>🔴 SALIENDO DESDE: {pista_A}</h3>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-vs {clase_win_A}'><p class='kpi-vs-title'>Costo Promedio x Hectárea</p><p class='kpi-vs-value'>$ {formato_latino(costo_ha_A, 0)}</p></div>", unsafe_allow_html=True)
        
        m_a1, m_a2 = st.columns(2)
        m_a1.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Costo Promedio x OS</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(costo_os_A, 0)}</p></div>", unsafe_allow_html=True)
        m_a2.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Costo Total Facturado</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(inv_A, 0)}</p></div>", unsafe_allow_html=True)
        
        st.info(f"**📊 Rendimiento:** {formato_latino(ha_A, 1)} Ha en {vuelos_A} Órdenes de Servicio.")

    with col_vs:
        st.markdown("<br><br><h1 style='text-align:center; color:#d4af37; font-size: 50px; font-family:Arial Black;'>VS</h1>", unsafe_allow_html=True)

    with col_B:
        st.markdown(f"<h3 style='text-align:center; color:#2F75B5;'>🔵 SALIENDO DESDE: {pista_B}</h3>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-vs {clase_win_B}'><p class='kpi-vs-title'>Costo Promedio x Hectárea</p><p class='kpi-vs-value'>$ {formato_latino(costo_ha_B, 0)}</p></div>", unsafe_allow_html=True)
        
        m_b1, m_b2 = st.columns(2)
        m_b1.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Costo Promedio x OS</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(costo_os_B, 0)}</p></div>", unsafe_allow_html=True)
        m_b2.markdown(f"<div class='kpi-vs' style='padding: 10px;'><p class='kpi-vs-title'>Costo Total Facturado</p><p class='kpi-vs-value' style='font-size:20px;'>$ {formato_latino(inv_B, 0)}</p></div>", unsafe_allow_html=True)

        st.info(f"**📊 Rendimiento:** {formato_latino(ha_B, 1)} Ha en {vuelos_B} Órdenes de Servicio.")

    st.markdown("<hr>", unsafe_allow_html=True)

    # --- GRÁFICAS DE COMPARACIÓN ---
    st.markdown("### 🛫 Rendimiento Logístico y Financiero")
    st.caption("Comparación visual de los costos según la pista de despegue.")

    df_ambos = pd.concat([df_A, df_B])
    
    if not df_ambos.empty:
        df_pistas = df_ambos.groupby('PISTA_MAESTRA').agg(
            COSTO_PROMEDIO_HA=('COSTO_HA_NUM', 'mean'),
            COSTO_TOTAL=('COSTO_TOTAL_OS', 'sum'),
            VUELOS=('OS_MAESTRA', 'nunique')
        ).reset_index()
        
        # Calcular el Costo Promedio por OS para la gráfica
        df_pistas['COSTO_PROMEDIO_OS'] = df_pistas['COSTO_TOTAL'] / df_pistas['VUELOS']

        c_graf1, c_graf2 = st.columns(2)

        with c_graf1:
            fig_ha = px.bar(
                df_pistas, 
                x='PISTA_MAESTRA', 
                y='COSTO_PROMEDIO_HA', 
                color='PISTA_MAESTRA', 
                text='COSTO_PROMEDIO_HA',
                color_discrete_map={pista_A: '#dc3545', pista_B: '#2F75B5'},
                title="Costo Promedio por Hectárea"
            )
            fig_ha.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont=dict(family="Arial Black"))
            fig_ha.update_layout(yaxis_title="$ / Ha", xaxis_title="Pista", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='#ffffff', showlegend=False)
            st.plotly_chart(fig_ha, use_container_width=True)

        with c_graf2:
            fig_os = px.bar(
                df_pistas, 
                x='PISTA_MAESTRA', 
                y='COSTO_PROMEDIO_OS', 
                color='PISTA_MAESTRA', 
                text='COSTO_PROMEDIO_OS',
                color_discrete_map={pista_A: '#dc3545', pista_B: '#2F75B5'},
                title="Costo Promedio por Orden de Servicio (OS)"
            )
            fig_os.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont=dict(family="Arial Black"))
            fig_os.update_layout(yaxis_title="$ / OS", xaxis_title="Pista", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='#ffffff', showlegend=False)
            st.plotly_chart(fig_os, use_container_width=True)

    # --- CARRERA EN EL TIEMPO ---
    st.markdown("### 📈 Tendencia: Costo por Hectárea a lo largo del tiempo")
    
    if not df_ambos.empty:
        df_ambos['MES_AÑO'] = df_ambos['FECHA_DT'].dt.to_period('M').astype(str)
        df_tendencia = df_ambos.groupby(['MES_AÑO', 'PISTA_MAESTRA'])['COSTO_HA_NUM'].mean().reset_index()
        df_tendencia = df_tendencia.sort_values(by='MES_AÑO')

        fig_line = px.line(
            df_tendencia, 
            x='MES_AÑO', 
            y='COSTO_HA_NUM', 
            color='PISTA_MAESTRA',
            markers=True,
            color_discrete_map={pista_A: '#dc3545', pista_B: '#2F75B5'},
            title=f"Evolución del Costo para {finca_sel}"
        )
        fig_line.update_traces(line=dict(width=4), marker=dict(size=10))
        fig_line.update_layout(yaxis_title="Costo ($ / Ha)", xaxis_title="Mes de Operación", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='#ffffff', hovermode="x unified")
        st.plotly_chart(fig_line, use_container_width=True)

if __name__ == "__main__":
    pass
