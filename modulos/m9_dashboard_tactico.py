import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import io

# =================================================================
# ⚡ MOTORES DE CONEXIÓN PROPIO (INTERCONEXIÓN DIRECTA EN RAM)
# =================================================================

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread_propio():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

def acortar_fecha(txt):
    try: return txt.split('(')[1].replace(')','') + " '" + txt[2:4]
    except Exception: return txt

@st.cache_data(show_spinner=False)
def cargar_y_preprocesar_boveda_mando_directo(_procesar_fecha_pesada, _extraer_numero):
    gc = inicializar_cliente_gspread_propio()
    datos_brutos = []
    
    if gc:
        try:
            url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
            sh = gc.open_by_url(url_maestra)
            ws = sh.worksheet("TABLA 1")
            datos_brutos = ws.get_all_values()
        except Exception:
            datos_brutos = []
            
    columnas_obj = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA", "INC_2026", "LIMITE", "ALERTA", "VAR_PCT", "COSTO_TOTAL", "PAGO_AVION"]

    if (not datos_brutos or len(datos_brutos) <= 2) and 'supabase' in st.session_state:
        try:
            supabase_client = st.session_state['supabase']
            respuesta_cloud = supabase_client.table("sap_tabla_1_maestro").select("*").execute()
            if respuesta_cloud.data:
                datos_brutos_supa = []
                for row in respuesta_cloud.data:
                    row_upper = {str(k).upper().strip(): v for k, v in row.items()}
                    fila_estructurada = [row_upper.get(col, "") for col in columnas_obj]
                    datos_brutos_supa.append(fila_estructurada)
                if datos_brutos_supa:
                    datos_brutos = [[""] * 30] * 5 + [columnas_obj] + datos_brutos_supa
        except Exception:
            pass
    
    if not datos_brutos or len(datos_brutos) <= 2: return pd.DataFrame()
        
    idx_headers = 4
    for i in range(min(8, len(datos_brutos))):
        row_clean = [str(x).strip().upper() for x in datos_brutos[i]]
        if "Nº ORDEN" in row_clean or "FINCA" in row_clean or "VALOR A FACTURAR" in "".join(row_clean):
            idx_headers = i
            break
        
    filas_datos = datos_brutos[idx_headers + 1:]
    lista_limpia = [r[:30] + [""] * max(0, 30 - len(r)) for r in filas_datos]
        
    df = pd.DataFrame(lista_limpia, columns=columnas_obj)
    
    def sanear_valores_sap(val):
        if pd.isna(val) or val is None: return 0.0
        s = str(val).strip().upper().replace("$", "").replace("COP", "").replace(" ", "")
        s_clean = re.sub(r'[^\d\.,\-]', '', s)
        if not s_clean or s_clean == '-': return 0.0
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
        except Exception:
            return 0.0

    cols_monetarias = ['COSTO_HA', 'VALOR_FACTURAR', 'LIMITE', 'COSTO_TOTAL', 'COSTO_AVION', 'DOMINICAL_HA']
    for col in cols_monetarias:
        df[col] = df[col].apply(sanear_valores_sap)
        
    df['AREA_FUMIG'] = df['AREA_FUMIG'].apply(lambda x: _extraer_numero(x) if str(x).strip() != "" else 0.0)
    df['REND_HR'] = df['REND_HR'].apply(lambda x: _extraer_numero(x) if str(x).strip() != "" else 0.0)
        
    df['FECHA_DT'] = df['FECHA'].apply(_procesar_fecha_pesada)
    df = df.dropna(subset=['FECHA_DT'])
    
    if df.empty: return pd.DataFrame()
    
    df['FECHA_DT'] = pd.to_datetime(df['FECHA_DT'])
    df['AÑO'] = df['FECHA_DT'].dt.year.astype(int)
    df['TRIMESTRE'] = df['FECHA_DT'].dt.quarter.astype(int)
    df['MES_NUM'] = df['FECHA_DT'].dt.month.astype(int)
    
    meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
    df['MES_NOMBRE'] = df['MES_NUM'].map(meses_dict)
    
    return df[df['AREA_FUMIG'] > 0].reset_index(drop=True)

# =================================================================
# 👑 FUNCIONES DE FORMATO LATINO
# =================================================================

def formato_latino(numero, decimales=0):
    if pd.isna(numero) or numero == 0: return "0"
    if decimales == 0: texto_us = f"{numero:,.0f}"
    else: texto_us = f"{numero:,.{decimales}f}"
    return texto_us.replace(",", "X").replace(".", ",").replace("X", ".")

def formato_gerencial_latino(numero):
    if pd.isna(numero) or numero == 0: return "$ 0"
    if numero >= 1_000_000: return f"$ {numero / 1_000_000:,.1f} M".replace(".", "X").replace(",", ".").replace("X", ",")
    elif numero >= 1_000: return f"$ {numero / 1_000:,.0f} K".replace(",", ".")
    else: return f"$ {formato_latino(numero, 0)}"

# =================================================================
# 👑 INTERFAZ GRÁFICA Y SEGMENTACIÓN DE TABLEROS (HUD VIP)
# =================================================================

def ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada):
    VERDE_INTENSO = '#143521' 
    DORADO = '#d4af37'         
    PALETA_YOY = [VERDE_INTENSO, '#27AE60', '#2F75B5', '#E67E22'] 
    
    st.markdown(f"""
    <style>
    .titulo-principal {{ color: {VERDE_INTENSO}; border-bottom: 3px solid {DORADO}; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }}
    .hud-comando {{ background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid {DORADO}; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
    .hud-comando-item {{ text-align: center; flex: 1; }}
    .hud-comando-title {{ font-size: 11px; font-weight: bold; color: {DORADO}; text-transform: uppercase; margin:0; letter-spacing: 1px; }}
    .hud-comando-value {{ font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }}
    
    div[data-testid="stSelectbox"] > div,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {{
        background-color: #ffffff !important;
        border: 3px solid {VERDE_INTENSO} !important;
        border-radius: 6px !important;
    }}
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        background-color: transparent !important;
        border: none !important;
    }}
    div[data-testid="stSelectbox"] * {{
        color: #000000 !important;
        font-weight: bold !important;
    }}
    div[data-testid="stSelectbox"] label p {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }}
    div[data-testid="stPlotlyChart"] {{
        transition: transform 0.3s ease-in-out, box-shadow 0.3s ease-in-out !important;
        border-radius: 10px !important;
        padding: 5px !important;
        background-color: #ffffff !important;
    }}
    div[data-testid="stPlotlyChart"]:hover {{
        transform: scale(1.02) !important;
        box-shadow: 0px 10px 25px rgba(20, 53, 33, 0.2) !important;
        z-index: 999 !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Centro de Comando: Rendimiento y Finanzas</h1>", unsafe_allow_html=True)
    
    df_dash = cargar_y_preprocesar_boveda_mando_directo(procesar_fecha_pesada, extraer_numero)
    
    if df_dash.empty:
        st.warning("⚠️ Bóveda vacía o sin misiones transaccionales activas registradas en la TABLA 1 o Supabase Cloud.")
        return

    st.markdown("### 🎛️ Filtros de Operación y Tiempo")
    
    t1, t2, t3 = st.columns(3)
    años_disp = ["TODOS (Comparativa Anual)"] + sorted(df_dash['AÑO'].unique().tolist(), reverse=True)
    año_sel = t1.selectbox("📅 AÑO FISCAL", años_disp, index=0)
    
    trimestres = {"TODOS": 0, "Q1 (Ene-Mar)": 1, "Q2 (Abr-Jun)": 2, "Q3 (Jul-Sep)": 3, "Q4 (Oct-Dic)": 4}
    trim_sel = t2.selectbox("📊 TRIMESTRE", list(trimestres.keys()))

    meses_disp = ["TODOS", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    mes_sel = t3.selectbox("📆 MES", meses_disp)

    f1, f2, f3 = st.columns(3)
    fincas_disp = ["TODAS"] + sorted(df_dash['FINCA'].astype(str).unique().tolist())
    pilotos_disp = ["TODOS"] + sorted(df_dash['PILOTO'].astype(str).unique().tolist())
    hks_disp = ["TODAS"] + sorted(df_dash['HK'].astype(str).unique().tolist())
    
    # 🎯 MEJORA VIP: El índice arranca en 1 (La primera finca real) para evitar el empacho visual inicial
    idx_finca_defecto = 1 if len(fincas_disp) > 1 else 0
    finca_filtro = f1.selectbox("📍 FINCA", fincas_disp, index=idx_finca_defecto)
    piloto_filtro = f2.selectbox("👨‍✈️ PILOTO", pilotos_disp)
    hk_filtro = f3.selectbox("✈️ MATRÍCULA (HK)", hks_disp)

    df_filtrado = df_dash.copy()
    if año_sel != "TODOS (Comparativa Anual)": df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(año_sel)]
    if trimestres[trim_sel] != 0: df_filtrado = df_filtrado[df_filtrado['TRIMESTRE'] == trimestres[trim_sel]]
    if mes_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['MES_NOMBRE'] == mes_sel]
    if finca_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['FINCA'] == finca_filtro]
    if piloto_filtro != "TODOS": df_filtrado = df_filtrado[df_filtrado['PILOTO'] == piloto_filtro]
    if hk_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['HK'] == hk_filtro]

    meses_nom = {1:"01-Ene", 2:"02-Feb", 3:"03-Mar", 4:"04-Abr", 5:"05-May", 6:"06-Jun", 7:"07-Jul", 8:"08-Ago", 9:"09-Sep", 10:"10-Oct", 11:"11-Nov", 12:"12-Dic"}
    df_filtrado['MES'] = df_filtrado['MES_NUM'].map(meses_nom).fillna("Desconocido")

    total_area = df_filtrado.groupby('FINCA')['AREA_FUMIG'].max().sum()
    total_facturacion = float(df_filtrado['COSTO_TOTAL'].sum())
    total_dominical = float(df_filtrado['DOMINICAL_HA'].sum())
    
    st.markdown(f"""
    <div class="hud-comando">
        <div class="hud-comando-item">
            <p class="hud-comando-title">Área Consolidada del Periodo</p>
            <p class="hud-comando-value">✈️ {formato_latino(total_area, 2)} ha</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Facturación Bruta Sincronizada</p>
            <p class="hud-comando-value">💰 $ {formato_latino(total_facturacion, 0)}</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Recargos Dominicales Aplicados</p>
            <p class="hud-comando-value" style="color: {DORADO};">⚠️ $ {formato_latino(total_dominical, 0)}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    titulo_finca = f" ({finca_filtro})" if finca_filtro != "TODAS" else " (COMPOSICIÓN GLOBAL)"
    g1, g2 = st.columns(2)

    # -----------------------------------------------------
    # GRÁFICO 1: ÁREA ASPERJADA (GRÁFICO DE ÁREA SUAVE)
    # -----------------------------------------------------
    with g1:
        st.markdown(f"#### ✈️ ÁREA ASPERJADA POR MES — {titulo_finca}", unsafe_allow_html=True)
        df_area_chart = df_filtrado.groupby(['MES_NUM', 'MES_NOMBRE', 'AÑO'])['AREA_FUMIG'].sum().reset_index()
        df_area_chart = df_area_chart.sort_values(by=['AÑO', 'MES_NUM']) 
        df_area_chart['AÑO_STR'] = df_area_chart['AÑO'].astype(str)
        df_area_chart['ETIQUETA'] = df_area_chart['AREA_FUMIG'].apply(lambda x: f"{formato_latino(x, 1)} ha")
        
        # 💥 Transformado a gráfico de área/líneas con relleno y curvas suaves
        fig1 = px.area(
            df_area_chart, x='MES_NOMBRE', y='AREA_FUMIG', color='AÑO_STR', 
            text='ETIQUETA', color_discrete_sequence=PALETA_YOY, markers=True
        )
        fig1.update_traces(
            textposition='top center', textfont=dict(size=11, family="Arial Black"),
            line=dict(shape='spline', width=3), # Curvas suaves y elegantes
            fill='tozeroy' # Relleno translúcido
        )
        fig1.update_layout(
            xaxis_title="", yaxis_title="Hectáreas (ha)", 
            plot_bgcolor='rgba(0,0,0,0)', legend_title_text='', 
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), 
            margin=dict(t=50, b=20)
        )
        if not df_area_chart.empty:
            fig1.update_yaxes(range=[0, df_area_chart['AREA_FUMIG'].max() * 1.3]) 
        st.plotly_chart(fig1, use_container_width=True)

    # -----------------------------------------------------
    # GRÁFICO 2: FACTURACIÓN vs LÍMITE (SE MANTIENE BARRAS VS LÍNEA POR SER UMBRAL)
    # -----------------------------------------------------
    with g2:
        st.markdown(f"#### ⚖️ FACTURACIÓN/ha vs LÍMITE — {titulo_finca}", unsafe_allow_html=True)
        
        if finca_filtro == "TODAS":
            df_costo = df_filtrado.groupby(['AÑO', 'FINCA']).agg({'VALOR_FACTURAR': 'mean', 'LIMITE': 'max'}).reset_index()
            df_costo = df_costo.sort_values(by=['AÑO', 'FINCA'])
            df_costo['ETIQUETA_X'] = df_costo['FINCA'].apply(lambda x: str(x)[:12] + '..' if len(str(x)) > 12 else str(x))
            df_costo['HOVER_FACT'] = df_costo['VALOR_FACTURAR'].apply(lambda x: f"$ {formato_latino(x, 0)} COP")
            df_costo['HOVER_LIMITE'] = df_costo['LIMITE'].apply(lambda x: f"$ {formato_latino(x, 0)} COP")

            go_fig = go.Figure()
            años_presentes = sorted(df_costo['AÑO'].unique())
            
            for i, año_map in enumerate(años_presentes):
                df_año = df_costo[df_costo['AÑO'] == año_map]
                color_asignado = PALETA_YOY[i % len(PALETA_YOY)]
                custom_data_hover = np.stack((df_año['FINCA'], df_año['HOVER_FACT'], df_año['AÑO']), axis=-1)
                
                go_fig.add_trace(go.Bar(
                    x=df_año['ETIQUETA_X'], y=df_año['VALOR_FACTURAR'], name=f"Promedio ({año_map})", 
                    marker_color=color_asignado, customdata=custom_data_hover,
                    hovertemplate='<b>Año:</b> %{customdata[2]}<br><b>Finca:</b> %{customdata[0]}<br><b>Promedio Facturado:</b> %{customdata[1]}<extra></extra>'
                ))
                
            go_fig.add_trace(go.Scatter(
                x=df_costo['ETIQUETA_X'], y=df_costo['LIMITE'], name="Límite Autorizado",
                mode='markers', marker=dict(color='#ff0000', size=10, symbol='line-ew', line=dict(width=3, color='#ff0000')),
                customdata=df_costo['HOVER_LIMITE'], hovertext=df_costo['FINCA'],
                hovertemplate='<b>Finca:</b> %{hovertext}<br><b>Límite Fijo:</b> %{customdata}<extra></extra>'
            ))
            
        else:
            df_filtrado['MES_ORDEN'] = df_filtrado['AÑO'].astype(str) + "-" + df_filtrado['MES_NUM'].astype(str).str.zfill(2) + " (" + df_filtrado['MES_NOMBRE'] + ")"
            df_costo = df_filtrado.groupby(['AÑO', 'MES_ORDEN', 'COCTEL']).agg({'VALOR_FACTURAR': 'mean', 'LIMITE': 'max'}).reset_index()
            df_costo = df_costo.sort_values(by=['AÑO', 'MES_ORDEN'])
            
            limite_real = df_filtrado[df_filtrado['LIMITE'] > 0]['LIMITE'].max()
            if pd.isna(limite_real) or limite_real == 0: limite_real = 200000 
            df_costo['LIMITE'] = df_costo['LIMITE'].apply(lambda x: limite_real if x == 0 else x)
            
            df_costo['FECHA_CORTA'] = df_costo['MES_ORDEN'].apply(acortar_fecha)
            df_costo['COCTEL_CORTO'] = df_costo['COCTEL'].apply(lambda x: str(x)[:10] + '..' if len(str(x)) > 10 else str(x))
            df_costo['ETIQUETA_X'] = df_costo['COCTEL_CORTO'] + "<br>(" + df_costo['FECHA_CORTA'] + ")"
            
            df_costo['HOVER_FACT'] = df_costo['VALOR_FACTURAR'].apply(lambda x: f"$ {formato_latino(x, 0)} COP")
            df_costo['HOVER_LIMITE'] = df_costo['LIMITE'].apply(lambda x: f"$ {formato_latino(x, 0)} COP")

            go_fig = go.Figure()
            años_presentes = sorted(df_costo['AÑO'].unique())
            for i, año_map in enumerate(años_presentes):
                df_año = df_costo[df_costo['AÑO'] == año_map]
                color_asignado = PALETA_YOY[i % len(PALETA_YOY)]
                custom_data_hover = np.stack((df_año['COCTEL'], df_año['HOVER_FACT'], df_año['AÑO']), axis=-1)
                
                go_fig.add_trace(go.Bar(
                    x=df_año['ETIQUETA_X'], y=df_año['VALOR_FACTURAR'], name=f"Facturación ({año_map})", 
                    marker_color=color_asignado, customdata=custom_data_hover,
                    hovertemplate='<b>Año:</b> %{customdata[2]}<br><b>Cóctel:</b> %{customdata[0]}<br><b>Facturación:</b> %{customdata[1]}<extra></extra>'
                ))
                
            go_fig.add_trace(go.Scatter(
                x=df_costo['ETIQUETA_X'], y=df_costo['LIMITE'], name="Límite Finca",
                mode='lines', line=dict(color='#ff0000', width=2, dash='dot'),
                customdata=df_costo['HOVER_LIMITE'], hovertext=df_costo['COCTEL'],
                hovertemplate='<b>Límite Fijo:</b> %{customdata}<extra></extra>'
            ))
            
        if not df_costo.empty:
            max_y = df_costo['LIMITE'].max() if df_costo['LIMITE'].max() > df_costo['VALOR_FACTURAR'].max() else df_costo['VALOR_FACTURAR'].max()
            go_fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', 
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), 
                yaxis=dict(title="Valor ($ COP / ha)", rangemode='tozero', range=[0, max_y * 1.2]), 
                xaxis=dict(showticklabels=False, title="Vuelos (Pase el cursor para ver detalles)"),
                margin=dict(b=20, t=50)
            )
            st.plotly_chart(go_fig, use_container_width=True)
        else:
            st.info("Sin datos de facturación para graficar.")
        
    st.markdown("<br>", unsafe_allow_html=True); g3, g4 = st.columns(2)

    # -----------------------------------------------------
    # GRÁFICO 3: RENDIMIENTO/HORA (LOLLIPOP CHART - PIRULETA)
    # -----------------------------------------------------
    with g3:
        st.markdown(f"#### ⏱️ RENDIMIENTO/Hora — {titulo_finca}", unsafe_allow_html=True)
        df_rend = df_filtrado.groupby(['HK', 'SEMANA'])['REND_HR'].sum().reset_index()
        df_rend['HK'] = df_rend['HK'].astype(str).str.replace(".0", "", regex=False)
        
        df_rend['SEMANA_NUM'] = pd.to_numeric(df_rend['SEMANA'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0).astype(int)
        df_rend = df_rend.sort_values(by=['HK', 'SEMANA_NUM'], ascending=[True, True])
        
        df_rend['EJE_Y'] = df_rend['HK'] + " | S" + df_rend['SEMANA_NUM'].astype(str)
        df_rend['ETIQUETA'] = df_rend['REND_HR'].apply(lambda x: f"{formato_latino(x, 2)} Hr")
        altura_dinamica = max(400, len(df_rend) * 22)
        
        # 💥 Transformado a Lollipop Chart (Línea delgada + Punto)
        fig3 = go.Figure()
        
        # El "tallo" de la piruleta (barra muy delgada)
        fig3.add_trace(go.Bar(
            y=df_rend['EJE_Y'], x=df_rend['REND_HR'], orientation='h', 
            marker_color=VERDE_INTENSO, width=0.1, showlegend=False, hoverinfo='skip'
        ))
        
        # La "cabeza" de la piruleta (punto con texto)
        fig3.add_trace(go.Scatter(
            x=df_rend['REND_HR'], y=df_rend['EJE_Y'], mode='markers+text',
            marker=dict(color=VERDE_INTENSO, size=12, line=dict(color='white', width=2)),
            text=df_rend['ETIQUETA'], textposition='middle right', textfont=dict(size=11, family="Arial Black"),
            showlegend=False, hovertemplate="<b>%{y}</b><br>Rendimiento: %{x} Hr<extra></extra>"
        ))
        
        fig3.update_layout(height=altura_dinamica, yaxis_title="Matrícula | Semana", xaxis_title="Rendimiento (Horas)", plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=10, b=20, r=40))
        fig3.update_yaxes(type='category', autorange="reversed") 
        if not df_rend.empty:
            fig3.update_xaxes(range=[0, df_rend['REND_HR'].max() * 1.3]) # Espacio extra para que el texto no se corte
        st.plotly_chart(fig3, use_container_width=True)
        
    # -----------------------------------------------------
    # GRÁFICO 4: FACTURACIÓN MENSUAL (LÍNEAS DE TRAYECTORIA)
    # -----------------------------------------------------
    with g4:
        st.markdown(f"#### 💵 FACTURACIÓN MENSUAL BASE — {titulo_finca}", unsafe_allow_html=True)
        df_mes = df_filtrado.groupby(['MES_NUM', 'MES_NOMBRE', 'AÑO'])['COSTO_TOTAL'].sum().reset_index()
        df_mes = df_mes.sort_values(by=['AÑO', 'MES_NUM'])
        df_mes['AÑO_STR'] = df_mes['AÑO'].astype(str)
        df_mes['TEXTO_GERENCIAL'] = df_mes['COSTO_TOTAL'].apply(formato_gerencial_latino)
        
        # 💥 Transformado a gráfico de Líneas Suaves
        fig4 = px.line(
            df_mes, x='MES_NOMBRE', y='COSTO_TOTAL', color='AÑO_STR', 
            text='TEXTO_GERENCIAL', color_discrete_sequence=PALETA_YOY, markers=True
        )
        fig4.update_traces(
            textposition='top center', textfont=dict(size=11, family="Arial Black"),
            line=dict(shape='spline', width=3.5), marker=dict(size=8)
        )
        fig4.update_layout(
            xaxis_title="", yaxis_title="Total Facturado ($ COP)", 
            plot_bgcolor='rgba(0,0,0,0)', legend_title_text='', 
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), 
            margin=dict(t=50, b=20)
        )
        if not df_mes.empty:
            fig4.update_yaxes(range=[0, df_mes['COSTO_TOTAL'].max() * 1.3])
        st.plotly_chart(fig4, use_container_width=True)

    # -----------------------------------------------------
    # GRÁFICO 5: RASTREO FINANCIERO DE RECARGOS DOMINICALES
    # -----------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(f"#### ⚠️ RASTREO FINANCIERO DE RECARGOS DOMINICALES — {titulo_finca}", unsafe_allow_html=True)
    df_dom = df_filtrado[df_filtrado['DOMINICAL_HA'] > 0].copy()
    if not df_dom.empty:
        df_dom['SEMANA_NUM'] = pd.to_numeric(df_dom['SEMANA'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0).astype(int)
        df_dom = df_dom.groupby(['AÑO', 'MES_NOMBRE', 'SEMANA_NUM'])['DOMINICAL_HA'].sum().reset_index()
        df_dom = df_dom.sort_values(by=['AÑO', 'SEMANA_NUM'])
        df_dom['AÑO_STR'] = df_dom['AÑO'].astype(str)
        df_dom['EJE_X'] = "Sem " + df_dom['SEMANA_NUM'].astype(str) + " (" + df_dom['MES_NOMBRE'] + ")"
        
        df_dom['ETIQUETA_DOM'] = df_dom['DOMINICAL_HA'].apply(lambda x: f"$ {formato_latino(x, 0)}")
        
        fig5 = px.bar(
            df_dom, x='EJE_X', y='DOMINICAL_HA', color='AÑO_STR', 
            barmode='group', text='ETIQUETA_DOM', 
            color_discrete_sequence=PALETA_YOY, 
            category_orders={"AÑO_STR": ["2025", "2026", "2027"]}
        )
        fig5.update_traces(textposition='outside', textfont=dict(size=12, color='black', family="Arial Black"), cliponaxis=False)
        fig5.update_layout(
            xaxis_title="Semana Operativa", yaxis_title="Total Recargos ($ COP)", 
            plot_bgcolor='rgba(0,0,0,0)', legend_title_text='', bargap=0.1, bargroupgap=0.0, 
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), 
            margin=dict(t=50)
        )
        if not df_dom.empty:
            fig5.update_yaxes(range=[0, df_dom['DOMINICAL_HA'].max() * 1.35]) 
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.info("✅ Excelente: No hay recargos dominicales registrados en el periodo seleccionado.")

    st.markdown("---")
    buffer_rep = io.BytesIO()
    df_filtrado.drop(columns=['FECHA_DT'], errors='ignore').to_excel(buffer_rep, sheet_name='Reporte', index=False)
    st.download_button(
        label="📥 DESCARGAR REPORTE EN EXCEL", 
        data=buffer_rep.getvalue(), 
        file_name=f"Reporte_Hectareas_{datetime.now().strftime('%Y%m%d')}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True
    )                        

if __name__ == "__main__":
    pass
