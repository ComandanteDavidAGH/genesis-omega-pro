import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 CONEXIÓN Y EXTRACCIÓN DE DATOS (Aislado para Gerencia)
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

@st.cache_data(show_spinner=False, ttl=600)
def cargar_datos_gerenciales():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    
    try:
        boveda_act = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        datos_brutos = boveda_act.worksheet("TABLA 1").get_all_values()
        
        if len(datos_brutos) > 5:
            # La columna 'COSTO_HA' corresponde a la Columna T (COSTO AVIÓN $/ha)
            columnas_t1 = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA"]
            filas_limpias = [r + [""]*(len(columnas_t1) - len(r)) for r in datos_brutos[5:]]
            df = pd.DataFrame([r[:len(columnas_t1)] for r in filas_limpias], columns=columnas_t1)
            
            # Limpieza básica
            df['FINCA'] = df['FINCA'].astype(str).str.strip().str.upper()
            df = df[~df['FINCA'].isin(['', 'NAN', 'NONE'])]
            
            # Clasificación de Tecnología
            def clasificar_tec(row):
                texto_busqueda = f"{str(row.get('PILOTO',''))} {str(row.get('HK',''))} {str(row.get('MODELO',''))} {str(row.get('PISTA',''))}".upper()
                if 'DRON' in texto_busqueda or 'DR5' in texto_busqueda:
                    return 'DRONE'
                return 'AVIÓN'
            
            df['TECNOLOGIA'] = df.apply(clasificar_tec, axis=1)
            
            # 💥 EXTRACCIÓN DE LAS DOS MÉTRICAS
            df['COSTO_FINAL_HA'] = df['VALOR_FACTURAR'].apply(limpiar_dinero) # Total
            df['COSTO_VUELO_HA'] = df['COSTO_HA'].apply(limpiar_dinero) # Solo Vuelo (Columna T)
            
            df = df[df['COSTO_FINAL_HA'] > 0]
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"🚨 Error al conectar con la bóveda: {e}")
        return pd.DataFrame()

# =================================================================
# 👑 RENDERIZADO VISUAL: PANEL GERENCIAL
# =================================================================

def ejecutar():
    st.header("", anchor="inicio_modulo")

    st.markdown("""
    <style>
        .titulo-gerencia { color: #1a365d; font-family: 'Arial Black'; border-bottom: 4px solid #d4af37; padding-bottom: 10px; margin-bottom: 20px;}
        .kpi-card { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 20px; border-radius: 10px; color: white; box-shadow: 0px 4px 15px rgba(0,0,0,0.2); text-align: center;}
        .kpi-title { font-size: 14px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
        .kpi-value { font-size: 28px; font-family: 'Arial Black'; margin: 10px 0 0 0; }
        div[data-testid="stDataFrame"] { border: 2px solid #e0e0e0; border-radius: 8px; box-shadow: 0px 4px 10px rgba(0,0,0,0.05); }
        .stTabs [data-baseweb="tab-list"] { gap: 10px; }
        .stTabs [data-baseweb="tab"] { background-color: #f0f2f6; border-radius: 5px 5px 0px 0px; padding: 10px 20px; font-weight: bold; }
        .stTabs [aria-selected="true"] { background-color: #0d1b2a !important; color: white !important; border-bottom: 3px solid #d4af37 !important;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-gerencia'>📊 Inteligencia Financiera: Drone vs Avión</h1>", unsafe_allow_html=True)

    with st.spinner("Extrayendo métricas financieras de la Bóveda..."):
        df_base = cargar_datos_gerenciales()

    if df_base.empty:
        st.warning("⚠️ No hay suficientes datos financieros procesados en la TABLA 1 para generar el comparativo.")
        return

    # 💥 CREACIÓN DE LAS PESTAÑAS
    tab_total, tab_vuelo = st.tabs(["💰 COSTO TOTAL OPERACIÓN ($/Ha)", "✈️ COSTO EXCLUSIVO DE VUELO ($/Ha)"])

    # ==========================================================
    # PESTAÑA 1: COSTO TOTAL (Químicos + Vuelo + Servicio)
    # ==========================================================
    with tab_total:
        costo_promedio_dron = df_base[df_base['TECNOLOGIA'] == 'DRONE']['COSTO_FINAL_HA'].mean()
        costo_promedio_avion = df_base[df_base['TECNOLOGIA'] == 'AVIÓN']['COSTO_FINAL_HA'].mean()

        if pd.isna(costo_promedio_dron): costo_promedio_dron = 0
        if pd.isna(costo_promedio_avion): costo_promedio_avion = 0

        diferencia_global = costo_promedio_avion - costo_promedio_dron
        ahorro_pct_global = (diferencia_global / costo_promedio_avion * 100) if costo_promedio_avion > 0 else 0

        st.markdown("### 🎯 Visión Global (Promedios Históricos - Total Operación)")
        kpi1, kpi2, kpi3 = st.columns(3)

        with kpi1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-title'>🚁 Promedio Histórico Dron</p><p class='kpi-value'>$ {costo_promedio_dron:,.0f} <span style='font-size:16px; color:#a0aec0;'>/ Ha</span></p></div>", unsafe_allow_html=True)
        with kpi2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-title'>🛩️ Promedio Histórico Avión</p><p class='kpi-value'>$ {costo_promedio_avion:,.0f} <span style='font-size:16px; color:#a0aec0;'>/ Ha</span></p></div>", unsafe_allow_html=True)
        with kpi3:
            color_diff = "#27ae60" if diferencia_global > 0 else "#e53e3e"
            estado_texto = "Ahorro a favor del Dron" if diferencia_global > 0 else "Dron es más costoso"
            st.markdown(f"<div class='kpi-card' style='border-left-color: {color_diff};'><p class='kpi-title'>⚖️ {estado_texto}</p><p class='kpi-value' style='color:{color_diff};'>$ {abs(diferencia_global):,.0f} <span style='font-size:16px;'>/ Ha ({ahorro_pct_global:+.1f}%)</span></p></div>", unsafe_allow_html=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)

        st.markdown("### 📋 Matriz Comparativa: Costo Total por Finca")
        
        matriz_fincas = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_FINAL_HA', aggfunc='mean').reset_index()

        if 'AVIÓN' not in matriz_fincas.columns: matriz_fincas['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_fincas.columns: matriz_fincas['DRONE'] = np.nan

        matriz_comparativa = matriz_fincas.dropna(subset=['AVIÓN', 'DRONE']).copy()

        if not matriz_comparativa.empty:
            matriz_comparativa['Ahorro con Dron ($)'] = matriz_comparativa['AVIÓN'] - matriz_comparativa['DRONE']
            matriz_comparativa['Eficiencia (%)'] = (matriz_comparativa['Ahorro con Dron ($)'] / matriz_comparativa['AVIÓN']) * 100

            df_visual = matriz_comparativa.copy()
            df_visual['AVIÓN'] = df_visual['AVIÓN'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual['DRONE'] = df_visual['DRONE'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual['Ahorro con Dron ($)'] = df_visual['Ahorro con Dron ($)'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual['Eficiencia (%)'] = df_visual['Eficiencia (%)'].map("{:+.2f} %".format)

            df_visual = df_visual.sort_values(by='Eficiencia (%)', ascending=False)

            st.dataframe(df_visual, use_container_width=True, hide_index=True)
            
            st.markdown("<br>### 📈 Análisis Gráfico de Brechas (Total Operación)", unsafe_allow_html=True)
            
            matriz_grafico = matriz_comparativa.sort_values(by='AVIÓN', ascending=False).head(15).copy()
            matriz_grafico['FINCA_CORTA'] = matriz_grafico['FINCA'].apply(lambda x: x[:18] + '...' if len(x) > 18 else x)
            
            fig = go.Figure()
            fig.add_trace(go.Bar(x=matriz_grafico['FINCA_CORTA'], y=matriz_grafico['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig.add_trace(go.Bar(x=matriz_grafico['FINCA_CORTA'], y=matriz_grafico['DRONE'], name='Dron', marker_color='#d4af37'))

            fig.update_layout(
                title="Comparativo Costo TOTAL ($ COP / Ha) - Top 15 Fincas",
                xaxis_title="Finca", yaxis_title="Costo Total Promedio",
                barmode='group', plot_bgcolor='rgba(0,0,0,0)', hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                xaxis=dict(tickangle=-45)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("📌 En el historial actual no se detectan fincas que hayan sido fumigadas tanto con Dron como con Avión.")


    # ==========================================================
    # PESTAÑA 2: COSTO EXCLUSIVO DE VUELO (Columna T)
    # ==========================================================
    with tab_vuelo:
        # Aquí usamos la métrica pura de Vuelo
        vuelo_promedio_dron = df_base[df_base['TECNOLOGIA'] == 'DRONE']['COSTO_VUELO_HA'].mean()
        vuelo_promedio_avion = df_base[df_base['TECNOLOGIA'] == 'AVIÓN']['COSTO_VUELO_HA'].mean()

        if pd.isna(vuelo_promedio_dron): vuelo_promedio_dron = 0
        if pd.isna(vuelo_promedio_avion): vuelo_promedio_avion = 0

        diff_vuelo = vuelo_promedio_avion - vuelo_promedio_dron
        ahorro_pct_vuelo = (diff_vuelo / vuelo_promedio_avion * 100) if vuelo_promedio_avion > 0 else 0

        st.markdown("### 🎯 Visión Global (Promedios Históricos - Tarifa de Vuelo)")
        kv1, kv2, kv3 = st.columns(3)

        with kv1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-title'>🚁 Tarifa Promedio Dron</p><p class='kpi-value'>$ {vuelo_promedio_dron:,.0f} <span style='font-size:16px; color:#a0aec0;'>/ Ha</span></p></div>", unsafe_allow_html=True)
        with kv2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-title'>🛩️ Tarifa Promedio Avión</p><p class='kpi-value'>$ {vuelo_promedio_avion:,.0f} <span style='font-size:16px; color:#a0aec0;'>/ Ha</span></p></div>", unsafe_allow_html=True)
        with kv3:
            color_diff_v = "#27ae60" if diff_vuelo > 0 else "#e53e3e"
            estado_texto_v = "Ahorro a favor del Dron" if diff_vuelo > 0 else "Dron es más costoso"
            st.markdown(f"<div class='kpi-card' style='border-left-color: {color_diff_v};'><p class='kpi-title'>⚖️ {estado_texto_v}</p><p class='kpi-value' style='color:{color_diff_v};'>$ {abs(diff_vuelo):,.0f} <span style='font-size:16px;'>/ Ha ({ahorro_pct_vuelo:+.1f}%)</span></p></div>", unsafe_allow_html=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)

        st.markdown("### 📋 Matriz Comparativa: Tarifa Exclusiva de Vuelo por Finca")
        
        matriz_fincas_v = df_base.pivot_table(index='FINCA', columns='TECNOLOGIA', values='COSTO_VUELO_HA', aggfunc='mean').reset_index()

        if 'AVIÓN' not in matriz_fincas_v.columns: matriz_fincas_v['AVIÓN'] = np.nan
        if 'DRONE' not in matriz_fincas_v.columns: matriz_fincas_v['DRONE'] = np.nan

        matriz_comp_vuelo = matriz_fincas_v.dropna(subset=['AVIÓN', 'DRONE']).copy()

        if not matriz_comp_vuelo.empty:
            matriz_comp_vuelo['Ahorro con Dron ($)'] = matriz_comp_vuelo['AVIÓN'] - matriz_comp_vuelo['DRONE']
            matriz_comp_vuelo['Eficiencia (%)'] = (matriz_comp_vuelo['Ahorro con Dron ($)'] / matriz_comp_vuelo['AVIÓN']) * 100

            df_visual_v = matriz_comp_vuelo.copy()
            df_visual_v['AVIÓN'] = df_visual_v['AVIÓN'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual_v['DRONE'] = df_visual_v['DRONE'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual_v['Ahorro con Dron ($)'] = df_visual_v['Ahorro con Dron ($)'].map("$ {:,.0f}".format).str.replace(",", ".")
            df_visual_v['Eficiencia (%)'] = df_visual_v['Eficiencia (%)'].map("{:+.2f} %".format)

            df_visual_v = df_visual_v.sort_values(by='Eficiencia (%)', ascending=False)

            st.dataframe(df_visual_v, use_container_width=True, hide_index=True)
            
            st.markdown("<br>### 📈 Análisis Gráfico de Brechas (Solo Vuelo)", unsafe_allow_html=True)
            
            matriz_grafico_v = matriz_comp_vuelo.sort_values(by='AVIÓN', ascending=False).head(15).copy()
            matriz_grafico_v['FINCA_CORTA'] = matriz_grafico_v['FINCA'].apply(lambda x: x[:18] + '...' if len(x) > 18 else x)
            
            fig_v = go.Figure()
            fig_v.add_trace(go.Bar(x=matriz_grafico_v['FINCA_CORTA'], y=matriz_grafico_v['AVIÓN'], name='Avión', marker_color='#1a365d'))
            fig_v.add_trace(go.Bar(x=matriz_grafico_v['FINCA_CORTA'], y=matriz_grafico_v['DRONE'], name='Dron', marker_color='#d4af37'))

            fig_v.update_layout(
                title="Comparativo Tarifa VUELO ($ COP / Ha) - Top 15 Fincas",
                xaxis_title="Finca", yaxis_title="Tarifa Vuelo Promedio",
                barmode='group', plot_bgcolor='rgba(0,0,0,0)', hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                xaxis=dict(tickangle=-45)
            )
            st.plotly_chart(fig_v, use_container_width=True)
        else:
            st.info("📌 En el historial actual no se detectan fincas que hayan sido fumigadas tanto con Dron como con Avión.")
