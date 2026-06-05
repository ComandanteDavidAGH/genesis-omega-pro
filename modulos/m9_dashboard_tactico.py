import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# =================================================================
# ⚡ MOTORES DE CACHÉ Y COMPUTACIÓN ACELERADA (ALTA VELOCIDAD)
# =================================================================

@st.cache_data(show_spinner=False)
def cargar_y_preprocesar_boveda_mando(_descargar_matriz_rapida, _procesar_fecha_pesada, _extraer_numero):
    """ Descarga, estructura y decodifica el historial financiero una sola vez en RAM """
    url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    datos_brutos = _descargar_matriz_rapida(url_maestra, "TABLA 1")
    
    if len(datos_brutos) <= 5:
        return pd.DataFrame()
        
    columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "AREA_FUMIG", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "REND_HR", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_AVION", "COSTO_HA", "DOMINICAL_HA", "COSTO_FINCA", "VALOR_FACTURAR", "PISTA", "INC_2026", "LIMITE", "ALERTA", "VAR_PCT", "COSTO_TOTAL", "PAGO_AVION"]
    
    filas_limpias = [r + [""]*(len(columnas) - len(r)) for r in datos_brutos[5:]]
    df = pd.DataFrame([r[:len(columnas)] for r in filas_limpias], columns=columnas)
    
    cols_numericas = ['AREA_FUMIG', 'REND_HR', 'COSTO_HA', 'DOMINICAL_HA', 'VALOR_FACTURAR', 'LIMITE', 'COSTO_TOTAL', 'COSTO_AVION']
    for col in cols_numericas:
        df[col] = df[col].apply(_extraer_numero)
        
    df['FECHA_DT'] = df['FECHA'].apply(_procesar_fecha_pesada)
    df = df.dropna(subset=['FECHA_DT'])
    
    df['AÑO'] = df['FECHA_DT'].dt.year.astype(int)
    df['TRIMESTRE'] = df['FECHA_DT'].dt.quarter.astype(int)
    df['MES_NUM'] = df['FECHA_DT'].dt.month.astype(int)
    
    meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
    df['MES_NOMBRE'] = df['MES_NUM'].map(meses_dict)
    df['MES_ORDEN'] = df['AÑO'].astype(str) + "-" + df['MES_NUM'].astype(str).str.zfill(2) + " (" + df['MES_NOMBRE'] + ")"
    
    return df[df['AREA_FUMIG'] > 0].reset_index(drop=True)

# =================================================================
# 👑 INTERFAZ GRÁFICA Y SEGMENTACIÓN DE TABLEROS (HUD VIP)
# =================================================================

def ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    
    /* HUD de Mando Financiero Corporativo */
    .hud-comando {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .hud-comando-item { text-align: center; flex: 1; }
    .hud-comando-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-comando-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Centro de Comando: Rendimiento y Finanzas</h1>", unsafe_allow_html=True)
    
    # EXTRACCIÓN MAESTRA EN RAM CACHEADA (Elimina retardos de red síncronos)
    df_dash = cargar_y_preprocesar_boveda_mando(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero)
    
    if df_dash.empty:
        st.warning("⚠️ Bóveda vacía o sin misiones transaccionales activas registradas en la TABLA 1.")
        return

    # --- 🎛️ FILTROS TÁCTICOS AVANZADOS ---
    st.markdown("### 🎛️ Filtros de Operación (Incluye Meses)")
    
    t1, t2 = st.columns(2)
    años_disp = ["TODOS"] + sorted(df_dash['AÑO'].unique().tolist(), reverse=True)
    año_sel = t1.selectbox("📅 AÑO FISCAL", años_disp, index=0)
    
    opciones_periodo = [
        "TODOS", 
        "Q1 (Ene-Mar)", "Q2 (Abr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dic)",
        "Ene", "Feb", "Mar", "Abr", "May", "Jun", 
        "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
    ]
    periodo_sel = t2.selectbox("📊 PERIODO (Trimestre/Mes)", opciones_periodo)

    f1, f2, f3 = st.columns(3)
    fincas_disp = ["TODAS"] + sorted(df_dash['FINCA'].astype(str).unique().tolist())
    pilotos_disp = ["TODOS"] + sorted(df_dash['PILOTO'].astype(str).unique().tolist())
    hks_disp = ["TODAS"] + sorted(df_dash['HK'].astype(str).unique().tolist())
    
    finca_filtro = f1.selectbox("📍 FINCA", fincas_disp)
    piloto_filtro = f2.selectbox("👨‍✈️ PILOTO", pilotos_disp)
    hk_filtro = f3.selectbox("✈️ MATRÍCULA (HK)", hks_disp)

    # 🎯 MATEMÁTICA DEL FILTRADO EN LA RAM PARA TRIMESTRE O MES
    df_filtrado = df_dash.copy()
    if año_sel != "TODOS": 
        df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(año_sel)]
        
    if periodo_sel != "TODOS":
        if periodo_sel.startswith("Q"):
            trim_map = {"Q1 (Ene-Mar)": 1, "Q2 (Abr-Jun)": 2, "Q3 (Jul-Sep)": 3, "Q4 (Oct-Dic)": 4}
            df_filtrado = df_filtrado[df_filtrado['TRIMESTRE'] == trim_map[periodo_sel]]
        else:
            df_filtrado = df_filtrado[df_filtrado['MES_NOMBRE'] == periodo_sel]
            
    if finca_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['FINCA'] == finca_filtro]
    if piloto_filtro != "TODOS": df_filtrado = df_filtrado[df_filtrado['PILOTO'] == piloto_filtro]
    if hk_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['HK'] == hk_filtro]

    # --- 🏆 HUD DE TARJETAS DE MANDO (KPIs) ---
    total_area = df_filtrado['AREA_FUMIG'].max() if not df_filtrado.empty else 0.0
    
    total_facturacion = float(df_filtrado['COSTO_TOTAL'].sum())
    total_dominical = float(df_filtrado['DOMINICAL_HA'].sum())
    
    st.markdown(f"""
    <div class="hud-comando">
        <div class="hud-comando-item">
            <p class="hud-comando-title">Área Consolidada del Periodo</p>
            <p class="hud-comando-value">🚜 {total_area:,.2f} Ha</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Facturación Bruta Sincronizada</p>
            <p class="hud-comando-value">💰 $ {total_facturacion:,.0f}</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Recargos Dominicales Aplicados</p>
            <p class="hud-comando-value" style="color: #ff3333;">⚠️ $ {total_dominical:,.0f}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    
    if df_filtrado.empty:
        st.warning("⚠️ El Escuadrón no registró operaciones financieras con los filtros actuales seleccionados.")
    else:
        g1, g2 = st.columns(2)

        with g1:
            st.markdown("<h4 style='text-align:center;'>🚜 ÁREA ASPERJADA POR MES</h4>", unsafe_allow_html=True)
            df_area_chart = df_filtrado.groupby('MES_ORDEN')['AREA_FUMIG'].sum().reset_index().sort_values(by='MES_ORDEN')
            
            fig1 = px.bar(df_area_chart, x='MES_ORDEN', y='AREA_FUMIG', text='AREA_FUMIG', color_discrete_sequence=['#548235'])
            fig1.update_traces(texttemplate='%{text:.1f}', textposition='outside', textfont_size=13)
            fig1.update_layout(xaxis_title="Mes Operativo", yaxis_title="Hectáreas", plot_bgcolor='rgba(0,0,0,0)', uniformtext_minsize=11)
            st.plotly_chart(fig1, use_container_width=True)

        with g2:
            st.markdown("<h4 style='text-align:center;'>⚖️ FACTURACIÓN/ha vs LÍMITE COMPUESTO</h4>", unsafe_allow_html=True)
            
            df_costo = df_filtrado.groupby(['MES_ORDEN', 'COCTEL']).agg({
                'VALOR_FACTURAR': 'mean', 
                'LIMITE': 'max'
            }).reset_index()
            
            limite_real = df_filtrado[df_filtrado['LIMITE'] > 0]['LIMITE'].max()
            if pd.isna(limite_real) or limite_real == 0: limite_real = 200000 
                
            df_costo['LIMITE'] = df_costo['LIMITE'].apply(lambda x: limite_real if x == 0 else x)
            
            def acortar_fecha(txt):
                try: return txt.split('(')[1].replace(')','') + " '" + txt[2:4]
                except: return txt
                
            df_costo['FECHA_CORTA'] = df_costo['MES_ORDEN'].apply(acortar_fecha)
            df_costo['COCTEL_CORTO'] = df_costo['COCTEL'].apply(lambda x: str(x)[:10] + '..' if len(str(x)) > 10 else str(x))
            df_costo['ETIQUETA'] = df_costo['COCTEL_CORTO'] + "<br>(" + df_costo['FECHA_CORTA'] + ")"

            go_fig = go.Figure()
            go_fig.add_trace(go.Bar(
                x=df_costo['ETIQUETA'], y=df_costo['VALOR_FACTURAR'], name="Facturación/ha",
                marker_color='#548235', text=df_costo['VALOR_FACTURAR'], 
                texttemplate='$ %{text:,.0f}', textposition='outside', textfont=dict(size=11),
                hovertext=df_costo['COCTEL'], hovertemplate='<b>Cóctel:</b> %{hovertext}<br><b>Facturación:</b> $ %{y:,.0f} COP<extra></extra>'
            ))
            
            go_fig.add_trace(go.Scatter(
                x=df_costo['ETIQUETA'], y=df_costo['LIMITE'], name="Límite Finca",
                mode='lines+markers', line=dict(color='red', width=3), marker=dict(size=8),
                hovertemplate='<b>Límite Fijo:</b> $ %{y:,.0f} COP<extra></extra>'
            ))
            
            go_fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', 
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                yaxis=dict(title="Valor ($ COP / Ha)", rangemode='tozero'), margin=dict(b=100)
            )
            go_fig.update_xaxes(tickangle=-90, tickfont=dict(size=10)) 
            st.plotly_chart(go_fig, use_container_width=True)
            
        st.markdown("<br>", unsafe_allow_html=True); g3, g4 = st.columns(2)

        with g3:
            titulo_finca = f" {finca_filtro}" if finca_filtro != "TODAS" else ""
            st.markdown(f"<h4 style='text-align:center;'>⏱️ RENDIMIENTO/Hora FINCA{titulo_finca}</h4>", unsafe_allow_html=True)
            
            df_rend = df_filtrado.groupby(['HK', 'SEMANA'])['REND_HR'].sum().reset_index()
            df_rend['HK'] = df_rend['HK'].astype(str).str.replace(".0", "", regex=False)
            df_rend['SEMANA'] = df_rend['SEMANA'].astype(str).str.replace(".0", "", regex=False)
            df_rend['EJE_Y'] = df_rend['HK'] + " | Sem " + df_rend['SEMANA']
            df_rend = df_rend.sort_values(by=['HK', 'SEMANA'], ascending=[True, False])
            
            fig3 = px.bar(df_rend, y='EJE_Y', x='REND_HR', orientation='h', text='REND_HR', color_discrete_sequence=['#548235'])
            fig3.update_traces(texttemplate='%{text:.2f}', textposition='outside', textfont_size=13)
            fig3.update_layout(yaxis_title="Matrícula (HK) | Semana", xaxis_title="Rendimiento (Horas)", plot_bgcolor='rgba(0,0,0,0)')
            fig3.update_yaxes(type='category')
            st.plotly_chart(fig3, use_container_width=True)
            
        with g4:
            st.markdown("<h4 style='text-align:center;'>💵 FACTURACIÓN MENSUAL CONSOLIDADA</h4>", unsafe_allow_html=True)
            df_mes = df_filtrado.groupby('MES_ORDEN')['COSTO_TOTAL'].sum().reset_index().sort_values(by='MES_ORDEN')
            
            fig4 = px.bar(df_mes, x='MES_ORDEN', y='COSTO_TOTAL', text='COSTO_TOTAL', color_discrete_sequence=['#548235'])
            fig4.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont_size=13)
            fig4.update_layout(xaxis_title="Mes Operativo", yaxis_title="Total Facturado ($)", plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig4, use_container_width=True)
