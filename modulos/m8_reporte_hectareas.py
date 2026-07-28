import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import io
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =================================================================
# 🚁 RADAR DE HECTÁREAS - OMEGA V22 (CONEXIÓN DIRECTA A EXCEL/SHEETS)
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    
    # 🌟 RESTAURACIÓN DEL TÍTULO PRINCIPAL CON SELLO DE VERIFICACIÓN
    st.markdown("<h1 style='color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: \"Arial Black\", sans-serif; text-transform: uppercase;'>Radar de Hectáreas y Rendimiento <span style='font-size: 16px; color: #d4af37;'>[V22]</span></h1>", unsafe_allow_html=True)
    
    # 🚀 REFORZAMIENTO ESTÉTICO VIP AISLADO Y EFECTO LUPA (SC)
    st.markdown("""
    <style>
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    
    /* 💥 CONTROLES ENDURECIDOS: Forzar visibilidad extrema en radios, selectores y calendarios */
    div[data-testid="stMainBlockContainer"] div[data-testid="stTextInput"] input, 
    div[data-testid="stMainBlockContainer"] div[data-testid="stNumberInput"] input,
    div[data-testid="stMainBlockContainer"] div[data-testid="stSelectbox"] [data-baseweb="select"],
    div[data-testid="stMainBlockContainer"] div[data-testid="stDateInput"] input {
        border: 2px solid #0d1b2a !important;
        border-radius: 6px !important;
        background-color: #ffffff !important;
        color: #0d1b2a !important;
        font-weight: 800 !important;
        font-size: 15px !important;
    }
    
    /* Acentuación de contraste para las etiquetas st.radio */
    div[data-testid="stMainBlockContainer"] div[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p {
        color: #0d1b2a !important;
        font-weight: 800 !important;
    }

    /* 🔍 EFECTO LUPA (SC) PARA TARJETAS KPI */
    .kpi-card {
        background-color: #0d1b2a; 
        color: white; 
        padding: 20px; 
        border-radius: 10px; 
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        border: 1px solid #1a365d;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .kpi-card:hover {
        transform: translateY(-5px) scale(1.03);
        box-shadow: 0 12px 20px rgba(212, 175, 55, 0.3); /* Resplandor dorado */
        border: 1px solid #d4af37;
    }
    .kpi-title {
        margin: 0; 
        color: #d4af37; 
        font-size: 16px; 
        font-weight: bold; 
        text-transform: uppercase;
    }
    .kpi-value {
        margin: 10px 0 0 0; 
        font-size: 32px; 
        font-weight: 900;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # ====================================================================
    # 💥 DESTRUCTOR DE MEMORIA ABSOLUTO
    # ====================================================================
    col_vacia, col_sync = st.columns([3, 1])
    if col_sync.button("🔄 Sincronizar Datos", type="primary", use_container_width=True, key="btn_sync_m8"):
        st.cache_data.clear()
        if 'm8_datos_crudos' in st.session_state:
            del st.session_state['m8_datos_crudos']
        st.toast("✅ Memoria vieja destruida. Descargando datos frescos...", icon="🔄")
        st.rerun()
    st.markdown("---")

    def extraer_numero(val):
        if pd.isna(val) or val is None or str(val).strip() == "": return 0.0
        try:
            texto = str(val).upper().replace("$", "").replace("COP", "").strip()
            if "," in texto and "." in texto: texto = texto.replace(".", "").replace(",", ".")
            elif "," in texto: texto = texto.replace(",", ".")
            return float(texto.replace(" ", ""))
        except: return 0.0

    def procesar_fecha_pesada(val):
        if pd.isna(val) or val is None or str(val).strip() == "": return None
        texto = str(val).strip().split(" ")[0]
        if texto.isdigit():
            try: return (pd.to_datetime('1899-12-30') + pd.to_timedelta(int(texto), unit='D')).date()
            except: pass
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y'):
            try: return datetime.strptime(texto, fmt).date()
            except: pass
        return None

    def fmt_latino(val, decimales=2):
        """FUERZA EL FORMATO COLOMBIANO: Puntos en miles, comas en decimales"""
        try: return f"{float(val):,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except: return str(val) if val is not None else ""

    if descargar_matriz_rapida is None:
        st.error("🚨 Error técnico: No se detecta el motor de lectura de Google Sheets.")
        return

    if 'm8_datos_crudos' not in st.session_state:
        datos_dict = []
        with st.spinner("🛰️ Conectando al satélite de Google Sheets..."):
            try:
                url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                filas_gspread = descargar_matriz_rapida(url_maestra, "TABLA 1")
                
                if filas_gspread:
                    idx_header = 4
                    for i in range(min(12, len(filas_gspread))):
                        if "FINCA" in [str(x).upper().strip() for x in filas_gspread[i]]:
                            idx_header = i
                            break
                    
                    if len(filas_gspread) > idx_header:
                        columnas = [str(c).strip().upper() for c in filas_gspread[idx_header]]
                        for fila in filas_gspread[idx_header+1:]:
                            if len(fila) < len(columnas):
                                fila = fila + [""] * (len(columnas) - len(fila))
                            elif len(fila) > len(columnas):
                                fila = fila[:len(columnas)]
                            datos_dict.append(dict(zip(columnas, fila)))
            except Exception:
                datos_dict = []

        if not datos_dict and supabase_client is not None:
            try:
                respuesta_cloud = supabase_client.table("sap_tabla_1_maestro").select("*").execute()
                if respuesta_cloud.data:
                    datos_dict = [{str(k).upper().strip(): v for k, v in row.items()} for row in respuesta_cloud.data]
            except Exception:
                pass

        st.session_state['m8_datos_crudos'] = datos_dict

    raw_data = st.session_state['m8_datos_crudos']

    try:
        if not raw_data:
            st.warning("⚠️ **Alerta de Radar:** No se pudieron procesar los registros de Google Sheets.")
            return

        datos_limpios = []
        for row in raw_data:
            r_norm = {str(k).replace("\n", " ").strip().upper(): (str(v).strip() if v is not None else "") for k, v in row.items()}
            
            llave_ha = next((k for k in r_norm.keys() if "FUMIG" in k or "HA_NETAS" in k or "HECTAREAS" in k), None)
            llave_hr = next((k for k in r_norm.keys() if "RENDIMIENTO" in k or "HORAS" in k or "HOROMETRO" in k), None)
            llave_sem = next((k for k in r_norm.keys() if k in ["SEM", "SEMANA"]), None)
            
            f_dt = procesar_fecha_pesada(r_norm.get("FECHA", r_norm.get("FECHA_OPERACION", "")))
            if f_dt is None:
                continue

            datos_limpios.append({
                "PISTA": r_norm.get("PISTA", "").strip().upper(),
                "HK": r_norm.get("HK", r_norm.get("MATRICULA", "")).strip().upper(),
                "MODELO": r_norm.get("MODELO", "").strip().upper(),
                "FECHA_REAL": f_dt,
                "SEMANA": r_norm.get(llave_sem, "") if llave_sem else "",
                "HA_NETAS": extraer_numero(r_norm.get(llave_ha, "0") if llave_ha else "0"),
                "H_PROPORCIONAL": extraer_numero(r_norm.get(llave_hr, "0") if llave_hr else "0")
            })

        df_rep = pd.DataFrame(datos_limpios)
        
        if df_rep.empty:
            st.warning("⚠️ No se encontraron registros con fechas procesables.")
            return

        mask_hk = df_rep['HK'] != ""
        mapa_modelo = {}
        if not df_rep[mask_hk].empty:
            mapa_flota = df_rep[mask_hk].groupby('HK')['PISTA'].agg(lambda x: x.value_counts().index[0] if not x.empty else "").to_dict()
            df_rep.loc[mask_hk, 'PISTA'] = df_rep.loc[mask_hk, 'HK'].map(mapa_flota).fillna(df_rep.loc[mask_hk, 'PISTA'])
            mapa_modelo = df_rep[mask_hk].groupby('HK')['MODELO'].first().to_dict()
        
        df_rep = df_rep[(df_rep['PISTA'] != "") & (df_rep['HA_NETAS'] > 0)]
        
        if df_rep.empty: return

        pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
        
        # --- 🎛️ PANEL DE CONTROL ---
        st.markdown("### 🎛️ Centro de Comando y Filtros")
        
        c1, c2, c3, c4 = st.columns([1.5, 1.0, 1.0, 1.2])
        vista_seleccionada = c1.radio("👁️ Vista Operativa:", ["📊 Resumen Gerencial", "📅 Mapa Semanal", "📈 Dashboard Ejecutivo"], horizontal=True, key="m8_v_final_v10")
        
        fecha_sel_ini = c2.date_input("📅 F. Inicial:", value=date(2026, 1, 1), min_value=date(2024, 1, 1), max_value=date(2030, 12, 31), key="m8_dat_ini_v6")
        fecha_sel_fin = c3.date_input("📅 F. Final:", value=date(2026, 12, 31), min_value=date(2024, 1, 1), max_value=date(2030, 12, 31), key="m8_dat_fin_v6")
        pista_sel = c4.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp, key="m8_pista_v6")

        if vista_seleccionada != "📈 Dashboard Ejecutivo":
            cc1, cc2, cc3 = st.columns(3)
            mostrar_horas = cc1.checkbox("⏱️ Mostrar Horas", value=True, key="m8_h_v6")
            calcular_rend_prom = cc2.checkbox("🚀 Mostrar Rend. (Ha/Hr)", value=True, key="m8_r_v6")
            agrupar_avion = cc3.toggle("✈️ Desglosar por Flota", value=False, key="m8_f_v6")

        df_filt = df_rep[(df_rep['FECHA_REAL'] >= fecha_sel_ini) & (df_rep['FECHA_REAL'] <= fecha_sel_fin)].copy()
        if pista_sel != "TODAS":
            df_filt = df_filt[df_filt['PISTA'] == pista_sel]
        
        if df_filt.empty:
            st.warning(f"⚠️ No hay registros de vuelo para {pista_sel} en el rango seleccionado.")
            return
            
        meses_nom = {1:"01-Ene", 2:"02-Feb", 3:"03-Mar", 4:"04-Abr", 5:"05-May", 6:"06-Jun", 7:"07-Jul", 8:"08-Ago", 9:"09-Sep", 10:"10-Oct", 11:"11-Nov", 12:"12-Dic"}
        df_filt['MES_NUM'] = df_filt['FECHA_REAL'].apply(lambda x: x.month)
        df_filt['MES'] = df_filt['MES_NUM'].apply(lambda x: meses_nom.get(x, "Desconocido"))
        
        st.markdown("---")
        rango_txt = f"{fecha_sel_ini.strftime('%d/%m/%Y')} al {fecha_sel_fin.strftime('%d/%m/%Y')}"
        
        # =================================================================
        # 📈 VISTA 3: DASHBOARD EJECUTIVO (KPIs COLOMBIANOS Y EFECTO LUPA)
        # =================================================================
        if vista_seleccionada == "📈 Dashboard Ejecutivo":
            st.markdown(f"#### 📈 Dashboard Ejecutivo y Participación Global ({rango_txt})")
            
            total_ha = df_filt['HA_NETAS'].sum()
            total_vuelos = len(df_filt)
            promedio_ha = total_ha/total_vuelos if total_vuelos>0 else 0
            
            df_dash = df_filt.groupby('PISTA').agg(
                VUELOS=('PISTA', 'count'),
                HECTAREAS=('HA_NETAS', 'sum')
            ).reset_index()
            
            df_dash['% VUELOS'] = (df_dash['VUELOS'] / total_vuelos) * 100
            df_dash['% HECTAREAS'] = (df_dash['HECTAREAS'] / total_ha) * 100
            df_dash = df_dash.sort_values(by='HECTAREAS', ascending=False)
            
            # FORMATO COLOMBIANO ESTRICTO EN KPIs
            ha_str = fmt_latino(total_ha, 2)
            vuelos_str = fmt_latino(total_vuelos, 0)
            prom_str = fmt_latino(promedio_ha, 2)
            
            k1, k2, k3 = st.columns(3)
            # EFECTO LUPA Y COLORES APLICADOS DESDE EL CSS DE ARRIBA
            k1.markdown(f"<div class='kpi-card'><p class='kpi-title'>Total Hectáreas</p><p class='kpi-value'>{ha_str}</p></div>", unsafe_allow_html=True)
            k2.markdown(f"<div class='kpi-card'><p class='kpi-title'>Total Misiones (Registros)</p><p class='kpi-value'>{vuelos_str}</p></div>", unsafe_allow_html=True)
            k3.markdown(f"<div class='kpi-card'><p class='kpi-title'>Promedio Ha/Misión</p><p class='kpi-value'>{prom_str}</p></div>", unsafe_allow_html=True)
            st.write("")

            g1, g2 = st.columns(2)
            df_dash['TXT_PCT'] = df_dash['% HECTAREAS'].apply(lambda x: f"{x:.1f}%".replace(".", ","))
            
            # Gráfico Dona
            fig_pie = px.pie(df_dash, values='VUELOS', names='PISTA', hole=0.45, 
                             title="<b>Distribución de Vuelos por Pista</b>", color_discrete_sequence=px.colors.qualitative.Prism)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label', texttemplate='%{label}<br>%{percent}')
            fig_pie.update_layout(separators=",.", showlegend=False, margin=dict(t=40, b=0, l=0, r=0))
            g1.plotly_chart(fig_pie, use_container_width=True)

            # Gráfico Barras
            fig_bar = px.bar(df_dash.sort_values('HECTAREAS', ascending=True), 
                             x='HECTAREAS', y='PISTA', orientation='h',
                             title="<b>Volumen de Hectáreas por Pista</b>",
                             text='TXT_PCT',
                             color='HECTAREAS', color_continuous_scale='Blues')
            fig_bar.update_layout(separators=",.", xaxis_title="Hectáreas Netas", yaxis_title="", coloraxis_showscale=False, margin=dict(t=40, b=0, l=0, r=0))
            g2.plotly_chart(fig_bar, use_container_width=True)

            st.markdown("##### 📋 Desglose de Participación Total")
            
            # 🛡️ SOLUCIÓN BLINDADA DEL ERROR "MISMATCHED COLUMNS": Se usa concat en lugar de loc
            df_tabla_dash = df_dash[['PISTA', 'VUELOS', 'HECTAREAS', '% VUELOS', '% HECTAREAS']].copy()
            fila_total = pd.DataFrame([{
                'PISTA': '👑 TOTAL GENERAL', 
                'VUELOS': df_tabla_dash['VUELOS'].sum(), 
                'HECTAREAS': df_tabla_dash['HECTAREAS'].sum(), 
                '% VUELOS': 100.0, 
                '% HECTAREAS': 100.0
            }])
            df_tabla_dash = pd.concat([df_tabla_dash, fila_total], ignore_index=True)
            
            fmt_dash = {
                "VUELOS": lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "HECTAREAS": lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "% VUELOS": lambda x: f"{x:.2f}%".replace(".", ","),
                "% HECTAREAS": lambda x: f"{x:.2f}%".replace(".", ",")
            }
            
            st.dataframe(
                df_tabla_dash.style.format(fmt_dash)
                  .bar(subset=["% HECTAREAS"], color='#d4af37', vmin=0, vmax=100)
                  .bar(subset=["% VUELOS"], color='#5c88b0', vmin=0, vmax=100),
                use_container_width=True, hide_index=True
            )

            st.markdown("---")
            st.markdown("##### 🗓️ Matriz de Participación Mensual (%)")
            
            matriz_ha = pd.pivot_table(df_filt, values='HA_NETAS', index='PISTA', columns='MES', aggfunc='sum', fill_value=0)
            cols_ordenadas = sorted(matriz_ha.columns, key=lambda x: int(str(x).split("-")[0]) if "-" in str(x) else 999)
            matriz_ha = matriz_ha[cols_ordenadas]
            
            matriz_pct = matriz_ha.div(matriz_ha.sum(axis=0), axis=1) * 100
            
            # 🚨 FUERZA BRUTA: El 100.0% exacto 
            totales_mes = [100.0 if matriz_ha[col].sum() > 0 else 0.0 for col in matriz_pct.columns]
            matriz_pct.loc['TOTAL MES'] = totales_mes
            
            st.dataframe(
                matriz_pct.style.format(lambda x: f"{x:.1f}%".replace(".", ","))
                .background_gradient(cmap="Blues", axis=None, subset=(matriz_pct.index[:-1], matriz_pct.columns)), 
                use_container_width=True
            )

        # =================================================================
        # 📊 VISTA 1 & 2: VISTAS CLÁSICAS
        # =================================================================
        elif vista_seleccionada == "📊 Resumen Gerencial":
            st.markdown(f"#### 📑 Consolidado Operativo ({rango_txt})")
            tabla_final = []
            total_hr_gral, total_ha_gral = 0, 0

            if agrupar_avion:
                df_gerencia = df_filt.groupby(['PISTA', 'HK', 'MES']).agg(REND_HR=('H_PROPORCIONAL', 'sum'), AREA_FUMIG=('HA_NETAS', 'sum')).reset_index()
                for pista in sorted(df_gerencia['PISTA'].unique()):
                    df_pista = df_gerencia[df_gerencia['PISTA'] == pista]
                    sum_hr_pista = df_pista['REND_HR'].sum()
                    sum_ha_pista = df_pista['AREA_FUMIG'].sum()
                    
                    fila_pista = {'NIVEL': f"📍 BASE: {pista}", 'AVIÓN (HK)': '', 'MES': 'TOTAL BASE'}
                    if mostrar_horas or calcular_rend_prom: fila_pista['REND (hr)'] = sum_hr_pista
                    fila_pista['ÁREA FUMIG (ha)'] = sum_ha_pista
                    if calcular_rend_prom: fila_pista['PROMEDIO (Ha/Hr)'] = sum_ha_pista / sum_hr_pista if sum_hr_pista > 0 else 0.0
                    tabla_final.append(fila_pista)
                    
                    for hk in sorted(df_pista['HK'].unique()):
                        datos_hk = df_pista[df_pista['HK'] == hk].sort_values(by='MES')
                        sum_hr_hk = datos_hk['REND_HR'].sum()
                        sum_ha_hk = datos_hk['AREA_FUMIG'].sum()
                        
                        modelo = str(mapa_modelo.get(hk, "")).upper()
                        es_dron = "DRON" in modelo or "DR" in hk
                        emoji = "🛸 DRON:" if es_dron else "✈️ AVION:"
                        
                        fila_hk = {'NIVEL': '', 'AVIÓN (HK)': f"{emoji} {hk}", 'MES': 'Total Flota'}
                        if mostrar_horas or calcular_rend_prom: fila_hk['REND (hr)'] = sum_hr_hk
                        fila_hk['ÁREA FUMIG (ha)'] = sum_ha_hk
                        if calcular_rend_prom: fila_hk['PROMEDIO (Ha/Hr)'] = sum_ha_hk / sum_hr_hk if sum_hr_hk > 0 else 0.0
                        tabla_final.append(fila_hk)
                        
                        for _, row in datos_hk.iterrows():
                            fila_mes = {'NIVEL': '', 'AVIÓN (HK)': '', 'MES': f"  ↳ {row['MES']}"}
                            if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = row['REND_HR']
                            fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                            if calcular_rend_prom: fila_mes['PROMEDIO (Ha/Hr)'] = row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0.0
                            tabla_final.append(fila_mes)
                            
                    total_hr_gral += sum_hr_pista
                    total_ha_gral += sum_ha_pista
                    
                fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'AVIÓN (HK)': '', 'MES': ''}
                if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = total_hr_gral
                fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                if calcular_rend_prom: fila_tot['PROMEDIO (Ha/Hr)'] = total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0.0
                tabla_final.append(fila_tot)
                
            else:
                df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(REND_HR=('H_PROPORCIONAL', 'sum'), AREA_FUMIG=('HA_NETAS', 'sum')).reset_index()
                for pista in sorted(df_gerencia['PISTA'].unique()):
                    datos_pista = df_gerencia[df_gerencia['PISTA'] == pista].sort_values(by='MES')
                    sum_hr = datos_pista['REND_HR'].sum()
                    sum_ha = datos_pista['AREA_FUMIG'].sum()
                    
                    fila_sub = {'NIVEL': f"📍 BASE: {pista}", 'MES': 'TOTAL BASE'}
                    if mostrar_horas or calcular_rend_prom: fila_sub['REND (hr)'] = sum_hr
                    fila_sub['ÁREA FUMIG (ha)'] = sum_ha
                    if calcular_rend_prom: fila_sub['PROMEDIO (Ha/Hr)'] = sum_ha / sum_hr if sum_hr > 0 else 0.0
                    tabla_final.append(fila_sub)
                    
                    for _, row in datos_pista.iterrows():
                        fila_mes = {'NIVEL': '', 'MES': f"  ↳ {row['MES']}"}
                        if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = row['REND_HR']
                        fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                        if calcular_rend_prom: fila_mes['PROMEDIO (Ha/Hr)'] = row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0.0
                        tabla_final.append(fila_mes)
                        
                    total_hr_gral += sum_hr
                    total_ha_gral += sum_ha
                    
                fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'MES': ''}
                if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = total_hr_gral
                fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                if calcular_rend_prom: fila_tot['PROMEDIO (Ha/Hr)'] = total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0.0
                tabla_final.append(fila_tot)

            df_visual = pd.DataFrame(tabla_final)
            
            def aplicar_estilos_originales(row):
                if "BASE:" in str(row['NIVEL']): return ['background-color: #d1ecf1; font-weight: bold; color: #0c5460;'] * len(row)
                elif "TOTAL GENERAL" in str(row['NIVEL']): return ['background-color: #c3e6cb; font-weight: bold; color: #155724;'] * len(row)
                elif 'AVIÓN (HK)' in row and ("✈️" in str(row.get('AVIÓN (HK)','')) or "🛸" in str(row.get('AVIÓN (HK)',''))):
                    return ['background-color: #f8f9fa; font-weight: bold; color: #212529;'] * len(row)
                return [''] * len(row)
                
            fmt_cols = {'ÁREA FUMIG (ha)': fmt_latino}
            if mostrar_horas or calcular_rend_prom: fmt_cols['REND (hr)'] = fmt_latino
            if calcular_rend_prom: fmt_cols['PROMEDIO (Ha/Hr)'] = fmt_latino
            
            st.dataframe(df_visual.style.apply(aplicar_estilos_originales, axis=1).format(fmt_cols), use_container_width=True, hide_index=True)

        else:
            matriz = pd.pivot_table(df_filt, values='HA_NETAS', index='MES', columns='SEMANA', aggfunc='sum', fill_value=0)
            matriz = matriz.sort_index()
            cols_ordenadas = sorted(matriz.columns, key=lambda x: int(x) if str(x).isdigit() else 999)
            matriz = matriz[cols_ordenadas]
            matriz['TOTAL MES'] = matriz.sum(axis=1)
            matriz.loc['TOTAL ANUAL'] = matriz.sum(axis=0)
            
            st.markdown(f"#### 🛩️ Rendimiento Semana a Semana ({rango_txt})")
            st.dataframe(matriz.style.format(fmt_latino).background_gradient(cmap="YlGn", axis=None), use_container_width=True)

        # =================================================================
        # 🎯 EXPORTACIÓN EXCEL GERENCIAL VIP CON ETIQUETAS
        # =================================================================
        st.markdown("---")
        buffer_rep = io.BytesIO()
        rango_label = f"{fecha_sel_ini.strftime('%Y%m%d')}_{fecha_sel_fin.strftime('%Y%m%d')}"
        
        if vista_seleccionada == "📈 Dashboard Ejecutivo":
            wb = Workbook()
            ws = wb.active
            ws.title = "Dashboard Ejecutivo"
            
            fill_header = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
            font_header = Font(color="FFFFFF", bold=True)
            fill_tot = PatternFill(start_color="D4AF37", end_color="D4AF37", fill_type="solid")
            font_tot = Font(color="000000", bold=True)
            borde = Border(left=Side(style='thin', color="CCCCCC"), right=Side(style='thin', color="CCCCCC"),
                           top=Side(style='thin', color="CCCCCC"), bottom=Side(style='thin', color="CCCCCC"))
            align_center = Alignment(horizontal='center', vertical='center')
            
            ws['B2'] = "REPORTE GERENCIAL: RADAR DE HECTÁREAS Y MISIONES"
            ws['B2'].font = Font(size=14, bold=True, color="0D1B2A")
            ws['B3'] = f"Período Analizado: {rango_txt}"
            ws['B3'].font = Font(italic=True, color="555555")
            
            df_export = df_dash.copy()
            total_vuelos = df_export['VUELOS'].sum()
            total_ha = df_export['HECTAREAS'].sum()
            
            headers = ['BASE OPERATIVA', 'TOTAL MISIONES', 'HECTÁREAS NETAS', '% MISIONES', '% HECTÁREAS']
            start_row = 6
            for col_idx, header in enumerate(headers, start=2):
                cell = ws.cell(row=start_row, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = borde
                
            curr_row = start_row + 1
            for _, row in df_export.iterrows():
                ws.cell(row=curr_row, column=2, value=row['PISTA']).border = borde
                ws.cell(row=curr_row, column=3, value=row['VUELOS']).number_format = '#,##0'
                ws.cell(row=curr_row, column=3).border = borde
                ws.cell(row=curr_row, column=4, value=row['HECTAREAS']).number_format = '#,##0.00'
                ws.cell(row=curr_row, column=4).border = borde
                ws.cell(row=curr_row, column=5, value=(row['VUELOS']/total_vuelos)).number_format = '0.00%'
                ws.cell(row=curr_row, column=5).border = borde
                ws.cell(row=curr_row, column=6, value=(row['HECTAREAS']/total_ha)).number_format = '0.00%'
                ws.cell(row=curr_row, column=6).border = borde
                curr_row += 1
                
            ws.cell(row=curr_row, column=2, value="TOTAL GENERAL").fill = fill_tot
            ws.cell(row=curr_row, column=2).font = font_tot
            ws.cell(row=curr_row, column=2).border = borde
            ws.cell(row=curr_row, column=3, value=total_vuelos).fill = fill_tot
            ws.cell(row=curr_row, column=3).font = font_tot
            ws.cell(row=curr_row, column=3).border = borde
            ws.cell(row=curr_row, column=3).number_format = '#,##0'
            ws.cell(row=curr_row, column=4, value=total_ha).fill = fill_tot
            ws.cell(row=curr_row, column=4).font = font_tot
            ws.cell(row=curr_row, column=4).border = borde
            ws.cell(row=curr_row, column=4).number_format = '#,##0.00'
            ws.cell(row=curr_row, column=5, value=1.0).fill = fill_tot
            ws.cell(row=curr_row, column=5).font = font_tot
            ws.cell(row=curr_row, column=5).border = borde
            ws.cell(row=curr_row, column=5).number_format = '0.00%'
            ws.cell(row=curr_row, column=6, value=1.0).fill = fill_tot
            ws.cell(row=curr_row, column=6).font = font_tot
            ws.cell(row=curr_row, column=6).border = borde
            ws.cell(row=curr_row, column=6).number_format = '0.00%'
            
            ws.column_dimensions['B'].width = 18
            ws.column_dimensions['C'].width = 18
            ws.column_dimensions['D'].width = 20
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 15
            
            # --- INCORPORACIÓN DE GRÁFICOS NATIVOS EXCEL CON ETIQUETAS VISIBLES ---
            data_len = len(df_export)
            cats = Reference(ws, min_col=2, min_row=start_row+1, max_row=start_row+data_len)
            
            bar_chart = BarChart()
            bar_chart.type = "bar"
            bar_chart.style = 11
            bar_chart.title = "Volumen de Hectáreas por Base"
            data_ha = Reference(ws, min_col=4, min_row=start_row, max_row=start_row+data_len)
            bar_chart.add_data(data_ha, titles_from_data=True)
            bar_chart.set_categories(cats)
            bar_chart.legend = None
            bar_chart.dataLabels = DataLabelList()
            bar_chart.dataLabels.showVal = True # ETIQUETAS EN BARRAS EXCEL
            ws.add_chart(bar_chart, "H5")
            
            pie_chart = DoughnutChart()
            pie_chart.title = "Distribución de Misiones (Vuelos)"
            pie_chart.style = 2
            data_vue = Reference(ws, min_col=3, min_row=start_row, max_row=start_row+data_len)
            pie_chart.add_data(data_vue, titles_from_data=True)
            pie_chart.set_categories(cats)
            pie_chart.dataLabels = DataLabelList()
            pie_chart.dataLabels.showPercent = True # ETIQUETAS EN DONA EXCEL
            pie_chart.dataLabels.showCatName = False
            ws.add_chart(pie_chart, "H20")
            
            wb.save(buffer_rep)

        else:
            df_export = df_visual if vista_seleccionada == "📊 Resumen Gerencial" else matriz
            df_export.to_excel(buffer_rep, sheet_name='Reporte', index=False if vista_seleccionada != "📅 Mapa Semanal" else True)
            
        st.download_button(
            label="💾 DESCARGAR REPORTE EJECUTIVO EN EXCEL",
            data=buffer_rep.getvalue(),
            file_name=f"Reporte_Ejecutivo_{rango_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )                        

    except Exception as e:
        st.error(f"🚨 Fallo procesando el reporte: {e}")

if __name__ == "__main__":
    pass
