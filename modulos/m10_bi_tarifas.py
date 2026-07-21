import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
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

# 💥 CACHÉ ACTIVO: MOTOR AUTÓNOMO M10
@st.cache_data(ttl=900, show_spinner=False)
def cargar_maestro_costos_m10(_procesar_fecha_pesada, _extraer_numero):
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
    
    patron_clean = re.compile(r'[^\d\.,\-]')
    
    def sanear_valores_sap(val):
        if pd.isna(val) or val is None or val == "": return 0.0
        s_clean = patron_clean.sub('', str(val).strip().upper().replace("$", "").replace("COP", "").replace(" ", ""))
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

    # PURIFICACIÓN FINANCIERA PROFUNDA
    cols_monetarias = ['COSTO_HA', 'VALOR_FACTURAR', 'LIMITE', 'COSTO_TOTAL', 'COSTO_AVION', 'COSTO_FINCA', 'PAGO_AVION']
    for col in cols_monetarias:
        df[col] = df[col].apply(sanear_valores_sap)
        
    df['AREA_FUMIG'] = df['AREA_FUMIG'].apply(lambda x: _extraer_numero(x) if str(x).strip() != "" else 0.0)
    
    df['FECHA_DT'] = df['FECHA'].apply(_procesar_fecha_pesada)
    df = df.dropna(subset=['FECHA_DT'])
    
    if df.empty: return pd.DataFrame()
    
    df['FECHA_DT'] = pd.to_datetime(df['FECHA_DT'])
    df['AÑO'] = df['FECHA_DT'].dt.year.astype(int)
    df['MES_NUM'] = df['FECHA_DT'].dt.month.astype(int)
    
    meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
    df['MES_NOMBRE'] = df['MES_NUM'].map(meses_dict)
    
    return df[df['AREA_FUMIG'] > 0].reset_index(drop=True)

# =================================================================
# 👑 FUNCIONES DE FORMATO
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
# 👑 INTERFAZ MÓDULO 10: INTELIGENCIA DE COSTOS (BI)
# =================================================================

def ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada):
    AZUL_PROFUNDO = '#0d1b2a' 
    DORADO = '#d4af37'         
    ROJO_ALERTA = '#c1121f'
    PALETA_COSTOS = ['#1d3557', '#457b9d', '#a8dadc', '#e63946']
    
    st.markdown(f"""
    <style>
    .titulo-principal {{ color: {AZUL_PROFUNDO}; border-bottom: 3px solid {DORADO}; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }}
    .hud-comando {{ background: linear-gradient(135deg, {AZUL_PROFUNDO} 0%, #1a365d 100%); border-left: 5px solid {DORADO}; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
    .hud-comando-item {{ text-align: center; flex: 1; border-right: 1px solid rgba(255,255,255,0.2); }}
    .hud-comando-item:last-child {{ border-right: none; }}
    .hud-comando-title {{ font-size: 11px; font-weight: bold; color: {DORADO}; text-transform: uppercase; margin:0; letter-spacing: 1px; }}
    .hud-comando-value {{ font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }}
    
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {{ border: 3px solid {AZUL_PROFUNDO} !important; border-radius: 6px !important; }}
    div[data-testid="stPlotlyChart"] {{ transition: transform 0.3s ease-in-out, box-shadow 0.3s ease-in-out !important; border-radius: 10px !important; padding: 5px !important; background-color: #ffffff !important; }}
    div[data-testid="stPlotlyChart"]:hover {{ transform: scale(1.04) !important; box-shadow: 0px 15px 30px rgba(13, 27, 42, 0.4), 0px 0px 15px rgba(212, 175, 55, 0.3) !important; z-index: 999 !important; }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Inteligencia de Costos y Rentabilidad (BI)</h1>", unsafe_allow_html=True)
    
    df_costos = cargar_maestro_costos_m10(procesar_fecha_pesada, extraer_numero)
    
    if df_costos.empty:
        st.warning("⚠️ Bóveda vacía o sin misiones transaccionales para procesar Inteligencia Financiera.")
        return

    st.markdown("### 🎛️ Filtros Financieros")
    
    t1, t2, t3 = st.columns(3)
    años_disp = ["TODOS"] + sorted(df_costos['AÑO'].unique().tolist(), reverse=True)
    año_sel = t1.selectbox("📅 AÑO FISCAL", años_disp, index=0)

    meses_disp = ["TODOS", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    mes_sel = t2.selectbox("📆 MES", meses_disp)

    fincas_disp = ["TODAS"] + sorted(df_costos['FINCA'].astype(str).unique().tolist())
    finca_filtro = t3.selectbox("📍 FINCA", fincas_disp)

    # --- PIPELINE DE FILTRADO ---
    df_filtrado = df_costos.copy()
    if año_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(año_sel)]
    if mes_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['MES_NOMBRE'] == mes_sel]
    if finca_filtro != "TODAS": df_filtrado = df_filtrado[df_filtrado['FINCA'] == finca_filtro]

    # --- CÁLCULO DE METRICAS FINANCIERAS MAESTRAS ---
    facturacion_bruta = df_filtrado['VALOR_FACTURAR'].sum()
    costo_total_operacion = df_filtrado['COSTO_TOTAL'].sum()
    rentabilidad_neta = facturacion_bruta - costo_total_operacion
    margen_pct = (rentabilidad_neta / facturacion_bruta * 100) if facturacion_bruta > 0 else 0
    costo_promedio_ha = df_filtrado['COSTO_HA'].mean()
    
    st.markdown(f"""
    <div class="hud-comando">
        <div class="hud-comando-item">
            <p class="hud-comando-title">Facturación Bruta</p>
            <p class="hud-comando-value" style="color: #4cc9f0;">$ {formato_gerencial_latino(facturacion_bruta)}</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Costos Operativos</p>
            <p class="hud-comando-value" style="color: #e63946;">$ {formato_gerencial_latino(costo_total_operacion)}</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Margen Bruto</p>
            <p class="hud-comando-value" style="color: {'#00ff66' if margen_pct >= 0 else '#ff3333'};">{formato_latino(margen_pct, 1)} %</p>
        </div>
        <div class="hud-comando-item">
            <p class="hud-comando-title">Costo Promedio / Ha</p>
            <p class="hud-comando-value">$ {formato_latino(costo_promedio_ha, 0)}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    # -----------------------------------------------------
    # GRÁFICO 1: ESTRUCTURA DE COSTOS (PIE)
    # -----------------------------------------------------
    with g1:
        st.markdown(f"#### 📊 Distribución del Costo Operativo", unsafe_allow_html=True)
        # Sumamos los dos pilares del costo
        costo_avion = df_filtrado['COSTO_AVION'].sum()
        costo_finca_insumos = df_filtrado['COSTO_FINCA'].sum()
        
        df_pie = pd.DataFrame({
            'Categoría': ['Costo Avión / Hora', 'Costo Insumos Finca'],
            'Valor': [costo_avion, costo_finca_insumos]
        })
        
        fig1 = px.pie(df_pie, values='Valor', names='Categoría', hole=0.4, color_discrete_sequence=['#1d3557', '#e63946'])
        fig1.update_traces(textposition='inside', textinfo='percent+label', hovertemplate="$ %{value:,.0f} COP")
        fig1.update_layout(plot_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=30, b=30))
        st.plotly_chart(fig1, use_container_width=True)

    # -----------------------------------------------------
    # GRÁFICO 2: RENTABILIDAD POR FINCA
    # -----------------------------------------------------
    with g2:
        st.markdown(f"#### ⚖️ Facturación vs Costo por Finca", unsafe_allow_html=True)
        df_rent = df_filtrado.groupby('FINCA').agg({'VALOR_FACTURAR': 'sum', 'COSTO_TOTAL': 'sum'}).reset_index()
        # Ordenar por los que más facturan
        df_rent = df_rent.sort_values(by='VALOR_FACTURAR', ascending=False).head(10)
        
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=df_rent['FINCA'], y=df_rent['VALOR_FACTURAR'], name='Facturación', marker_color='#4cc9f0'))
        fig2.add_trace(go.Bar(x=df_rent['FINCA'], y=df_rent['COSTO_TOTAL'], name='Costo', marker_color='#e63946'))
        
        fig2.update_layout(barmode='group', xaxis_tickangle=-45, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=30))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True); g3, g4 = st.columns(2)

    # -----------------------------------------------------
    # GRÁFICO 3: EVOLUCIÓN MENSUAL DEL MARGEN (LÍNEA)
    # -----------------------------------------------------
    with g3:
        st.markdown(f"#### 📈 Evolución del Margen Operativo (%)", unsafe_allow_html=True)
        df_evo = df_filtrado.groupby(['AÑO', 'MES_NUM', 'MES_NOMBRE']).agg({'VALOR_FACTURAR': 'sum', 'COSTO_TOTAL': 'sum'}).reset_index()
        df_evo = df_evo.sort_values(by=['AÑO', 'MES_NUM'])
        df_evo['MARGEN'] = ((df_evo['VALOR_FACTURAR'] - df_evo['COSTO_TOTAL']) / df_evo['VALOR_FACTURAR']) * 100
        df_evo['MARGEN'] = df_evo['MARGEN'].fillna(0)
        df_evo['EJE_X'] = df_evo['MES_NOMBRE'] + " " + df_evo['AÑO'].astype(str)
        df_evo['ETIQUETA'] = df_evo['MARGEN'].apply(lambda x: f"{formato_latino(x, 1)}%")

        fig3 = px.line(df_evo, x='EJE_X', y='MARGEN', text='ETIQUETA', markers=True, color_discrete_sequence=[DORADO])
        fig3.update_traces(textposition='top center', line=dict(width=4), marker=dict(size=10))
        fig3.update_layout(xaxis_title="", yaxis_title="Margen (%)", plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=30))
        st.plotly_chart(fig3, use_container_width=True)

    # -----------------------------------------------------
    # GRÁFICO 4: TOP 10 CÓCTELES MÁS COSTOSOS POR HECTÁREA
    # -----------------------------------------------------
    with g4:
        st.markdown(f"#### 🧪 Top 10 Cócteles más Costosos/Ha", unsafe_allow_html=True)
        df_coctel = df_filtrado.groupby('COCTEL')['COSTO_HA'].mean().reset_index()
        df_coctel = df_coctel.sort_values(by='COSTO_HA', ascending=True).tail(10) # Tail para que los más caros salgan arriba en Plotly
        df_coctel['COCTEL_CORTO'] = df_coctel['COCTEL'].apply(lambda x: str(x)[:20] + '..')
        df_coctel['ETIQUETA'] = df_coctel['COSTO_HA'].apply(lambda x: f"$ {formato_latino(x, 0)}")

        fig4 = px.bar(df_coctel, y='COCTEL_CORTO', x='COSTO_HA', orientation='h', text='ETIQUETA', color_discrete_sequence=['#457b9d'])
        fig4.update_traces(textposition='outside')
        fig4.update_layout(xaxis_title="Costo Promedio / Ha", yaxis_title="", plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=30))
        # Ajustamos el límite de X para que la etiqueta no se corte
        fig4.update_xaxes(range=[0, df_coctel['COSTO_HA'].max() * 1.3])
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")
    buffer_rep = io.BytesIO()
    df_filtrado.drop(columns=['FECHA_DT'], errors='ignore').to_excel(buffer_rep, sheet_name='Costos_BI', index=False)
    st.download_button(label="📥 DESCARGAR REPORTE DE COSTOS (EXCEL)", data=buffer_rep.getvalue(), file_name=f"Inteligencia_Costos_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)                        

if __name__ == "__main__":
    pass
