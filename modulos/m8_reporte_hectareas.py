import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# =================================================================
# 🚁 RADAR DE HECTÁREAS - MODO BLINDADO (ANTI-COLAPSOS)
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
    # 🟢 SEMÁFORO 1
    progreso = st.empty()
    progreso.info("🟢 FASE 1: Iniciando sistemas de radar...")

    def extraer_numero(val):
        if pd.isna(val) or str(val).strip() == "": return 0.0
        try:
            texto = str(val).upper().replace("$", "").replace("COP", "").strip()
            if "," in texto and "." in texto: texto = texto.replace(".", "").replace(",", ".")
            elif "," in texto: texto = texto.replace(",", ".")
            return float(texto.replace(" ", ""))
        except: return 0.0

    def procesar_fecha_pesada(val):
        if pd.isna(val) or str(val).strip() == "": return None
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%S.%f'):
            try: return datetime.strptime(str(val).strip().split(" ")[0], fmt.split(" ")[0])
            except: pass
        return None

    def fmt_latino(val):
        try: return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except: return str(val)

    if supabase_client is None:
        st.error("🚨 Supabase desconectado.")
        return

    try:
        # 🟢 SEMÁFORO 2
        progreso.info("🟢 FASE 2: Descargando bóveda de datos desde la nube...")
        respuesta = supabase_client.table("TABLA_1").select("*").limit(10000).execute()
        raw_data = respuesta.data
        
        if not raw_data:
            progreso.warning("⚠️ La TABLA_1 está vacía en Supabase.")
            return
            
        # 🟢 SEMÁFORO 3
        progreso.info(f"🟢 FASE 3: Limpiando {len(raw_data)} registros...")
        columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "HA_NETAS", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "H_PROPORCIONAL", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_TOTAL_AVION", "TARIFA_HA", "RECARGO_HA", "SUBTOTAL", "COSTO_HORA", "PISTA"]
        
        df_raw = pd.DataFrame(raw_data)
        df_raw.columns = [str(c).upper().strip() for c in df_raw.columns]
        
        for col in columnas:
            if col not in df_raw.columns:
                df_raw[col] = ""
                
        df_rep = df_raw[columnas].copy()
        df_rep['HA_NETAS'] = df_rep['HA_NETAS'].apply(extraer_numero)
        df_rep['H_PROPORCIONAL'] = df_rep['H_PROPORCIONAL'].apply(extraer_numero)
        df_rep['SEMANA'] = df_rep['SEMANA'].astype(str).str.strip()
        df_rep['PISTA'] = df_rep['PISTA'].astype(str).str.strip().str.upper()
        df_rep['HK'] = df_rep['HK'].astype(str).str.strip().str.upper()
        df_rep['MODELO'] = df_rep['MODELO'].astype(str).str.strip().str.upper()
        
        mask_hk = df_rep['HK'] != ""
        mapa_modelo = {}
        if not df_rep[mask_hk].empty:
            mapa_flota = df_rep[mask_hk].groupby('HK')['PISTA'].agg(lambda x: x.value_counts().index[0] if not x.empty else "").to_dict()
            df_rep.loc[mask_hk, 'PISTA'] = df_rep.loc[mask_hk, 'HK'].map(mapa_flota).fillna(df_rep.loc[mask_hk, 'PISTA'])
            mapa_modelo = df_rep[mask_hk].groupby('HK')['MODELO'].first().to_dict()
        
        df_rep = df_rep[(df_rep['PISTA'] != "") & (df_rep['HA_NETAS'] > 0)]
        df_rep['FECHA_DT'] = df_rep['FECHA'].apply(procesar_fecha_pesada)
        df_rep = df_rep.dropna(subset=['FECHA_DT'])
        
        if df_rep.empty:
            progreso.warning("⚠️ No hay fechas ni hectáreas válidas para armar el reporte.")
            return

        pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
        min_fecha_real = df_rep['FECHA_DT'].min().date()
        max_fecha_real = df_rep['FECHA_DT'].max().date()
        
        # 🟢 SEMÁFORO 4
        progreso.success("✅ Motores listos. Renderizando interfaz...")

        st.markdown("### 🎛️ Centro de Comando y Filtros")
        c1, c2, c3, c4 = st.columns([1.5, 1, 1, 1])
        
        vista_seleccionada = c1.radio("👁️ Seleccione la Vista del Radar:", ["📊 Resumen Gerencial", "📅 Mapa Semanal"], horizontal=True)
        fecha_sel_ini = c2.date_input("📅 Inicial:", value=min_fecha_real)
        fecha_sel_fin = c3.date_input("📅 Final:", value=max_fecha_real)
        pista_sel = c4.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp)
        
        mostrar_horas = False
        calcular_rend_prom = False
        agrupar_avion = False 

        if vista_seleccionada == "📊 Resumen Gerencial":
            cc1, cc2, cc3 = st.columns(3)
            mostrar_horas = cc1.checkbox("⏱️ Mostrar Horas", value=True)
            calcular_rend_prom = cc2.checkbox("🚀 Mostrar Rend. (Ha/Hr)", value=True)
            agrupar_avion = cc3.toggle("✈️ Desglosar por Flota", value=False)

        df_filt = df_rep[(df_rep['FECHA_DT'].dt.date >= fecha_sel_ini) & (df_rep['FECHA_DT'].dt.date <= fecha_sel_fin)].copy()
        if pista_sel != "TODAS":
            df_filt = df_filt[df_filt['PISTA'] == pista_sel]
        
        if df_filt.empty:
            st.warning("⚠️ No hay datos para estas fechas.")
            return
            
        meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
        df_filt['MES'] = df_filt['FECHA_DT'].dt.month.map(meses_nom)
        
        st.markdown("---")
        rango_txt = f"{fecha_sel_ini.strftime('%d/%m/%Y')} al {fecha_sel_fin.strftime('%d/%m/%Y')}"
        
        if vista_seleccionada == "📊 Resumen Gerencial":
            st.markdown(f"#### 📑 Consolidado ({rango_txt})")
            tabla_final = []
            total_hr_gral = 0
            total_ha_gral = 0

            if agrupar_avion:
                df_gerencia = df_filt.groupby(['PISTA', 'HK', 'MES']).agg(REND_HR=('H_PROPORCIONAL', 'sum'), AREA_FUMIG=('HA_NETAS', 'sum')).reset_index()
                for pista in sorted(df_gerencia['PISTA'].unique()):
                    df_pista = df_gerencia[df_gerencia['PISTA'] == pista]
                    sum_hr_pista = df_pista['REND_HR'].sum()
                    sum_ha_pista = df_pista['AREA_FUMIG'].sum()
                    
                    fila_pista = {'NIVEL': f"📍 BASE: {pista}", 'AVIÓN (HK)': '', 'MES': 'TOTAL BASE'}
                    if mostrar_horas or calcular_rend_prom: fila_pista['REND (hr)'] = fmt_latino(sum_hr_pista)
                    fila_pista['ÁREA FUMIG (ha)'] = fmt_latino(sum_ha_pista)
                    if calcular_rend_prom: fila_pista['PROMEDIO (Ha/Hr)'] = fmt_latino(sum_ha_pista / sum_hr_pista if sum_hr_pista > 0 else 0)
                    tabla_final.append(fila_pista)
                    
                    for hk in sorted(df_pista['HK'].unique()):
                        datos_hk = df_pista[df_pista['HK'] == hk].sort_values(by='MES')
                        sum_hr_hk = datos_hk['REND_HR'].sum()
                        sum_ha_hk = datos_hk['AREA_FUMIG'].sum()
                        
                        modelo = str(mapa_modelo.get(hk, "")).upper()
                        es_dron = "DRON" in modelo or "DR" in hk
                        emoji = "🚁 DRON:" if es_dron else "✈️ AVION:"
                        
                        fila_hk = {'NIVEL': '', 'AVIÓN (HK)': f"{emoji} {hk}", 'MES': 'Total Flota'}
                        if mostrar_horas or calcular_rend_prom: fila_hk['REND (hr)'] = fmt_latino(sum_hr_hk)
                        fila_hk['ÁREA FUMIG (ha)'] = fmt_latino(sum_ha_hk)
                        if calcular_rend_prom: fila_hk['PROMEDIO (Ha/Hr)'] = fmt_latino(sum_ha_hk / sum_hr_hk if sum_hr_hk > 0 else 0)
                        tabla_final.append(fila_hk)
                        
                        for _, row in datos_hk.iterrows():
                            fila_mes = {'NIVEL': '', 'AVIÓN (HK)': '', 'MES': row['MES']}
                            if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = fmt_latino(row['REND_HR'])
                            fila_mes['ÁREA FUMIG (ha)'] = fmt_latino(row['AREA_FUMIG'])
                            if calcular_rend_prom: fila_mes['PROMEDIO (Ha/Hr)'] = fmt_latino(row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0)
                            tabla_final.append(fila_mes)
                            
                    total_hr_gral += sum_hr_pista
                    total_ha_gral += sum_ha_pista
                    
                fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'AVIÓN (HK)': '', 'MES': ''}
                if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = fmt_latino(total_hr_gral)
                fila_tot['ÁREA FUMIG (ha)'] = fmt_latino(total_ha_gral)
                if calcular_rend_prom: fila_tot['PROMEDIO (Ha/Hr)'] = fmt_latino(total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0)
                tabla_final.append(fila_tot)
                
            else:
                df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(REND_HR=('H_PROPORCIONAL', 'sum'), AREA_FUMIG=('HA_NETAS', 'sum')).reset_index()
                for pista in sorted(df_gerencia['PISTA'].unique()):
                    datos_pista = df_gerencia[df_gerencia['PISTA'] == pista].sort_values(by='MES')
                    sum_hr = datos_pista['REND_HR'].sum()
                    sum_ha = datos_pista['AREA_FUMIG'].sum()
                    
                    fila_sub = {'NIVEL': f"📍 BASE: {pista}", 'MES': 'TOTAL BASE'}
                    if mostrar_horas or calcular_rend_prom: fila_sub['REND (hr)'] = fmt_latino(sum_hr)
                    fila_sub['ÁREA FUMIG (ha)'] = fmt_latino(sum_ha)
                    if calcular_rend_prom: fila_sub['PROMEDIO (Ha/Hr)'] = fmt_latino(sum_ha / sum_hr if sum_hr > 0 else 0)
                    tabla_final.append(fila_sub)
                    
                    for _, row in datos_pista.iterrows():
                        fila_mes = {'NIVEL': '', 'MES': row['MES']}
                        if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = fmt_latino(row['REND_HR'])
                        fila_mes['ÁREA FUMIG (ha)'] = fmt_latino(row['AREA_FUMIG'])
                        if calcular_rend_prom: fila_mes['PROMEDIO (Ha/Hr)'] = fmt_latino(row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0)
                        tabla_final.append(fila_mes)
                        
                    total_hr_gral += sum_hr
                    total_ha_gral += sum_ha
                    
                fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'MES': ''}
                if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = fmt_latino(total_hr_gral)
                fila_tot['ÁREA FUMIG (ha)'] = fmt_latino(total_ha_gral)
                if calcular_rend_prom: fila_tot['PROMEDIO (Ha/Hr)'] = fmt_latino(total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0)
                tabla_final.append(fila_tot)

            # 🎯 ARMADURA: Tabla plana sin estilos peligrosos
            df_visual = pd.DataFrame(tabla_final)
            st.dataframe(df_visual, use_container_width=True, hide_index=True)

        else:
            matriz = pd.pivot_table(df_filt, values='HA_NETAS', index='MES', columns='SEMANA', aggfunc='sum', fill_value=0)
            matriz.index = [m.split('-')[1] if '-' in m else m for m in matriz.index]
            matriz['TOTAL MES'] = matriz.sum(axis=1)
            matriz.loc['TOTAL ANUAL'] = matriz.sum(axis=0)
            
            st.markdown(f"#### 🛩️ Rendimiento Semana a Semana ({rango_txt})")
            st.dataframe(matriz.applymap(fmt_latino), use_container_width=True)
            
            df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
            if not df_grafico.empty:
                df_grafico['TXT'] = df_grafico['TOTAL MES'].apply(fmt_latino)
                fig = px.bar(df_grafico, x='index', y='TOTAL MES', text='TXT', color='TOTAL MES', color_continuous_scale='Greens')
                fig.update_traces(textposition='outside')
                fig.update_layout(xaxis_title="Mes", showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        # 🎯 ARMADURA EXCEL: Descarga limpia sin gráficos internos
        st.markdown("---")
        buffer_rep = io.BytesIO()
        nombre_hoja = 'Reporte'
        if vista_seleccionada == "📊 Resumen Gerencial":
            df_visual.to_excel(buffer_rep, sheet_name=nombre_hoja, index=False)
        else:
            matriz.to_excel(buffer_rep, sheet_name=nombre_hoja)
            
        rango_label = f"{fecha_sel_ini.strftime('%Y%m%d')}_{fecha_sel_fin.strftime('%Y%m%d')}"
        st.download_button(
            label="💾 DESCARGAR REPORTE EN EXCEL",
            data=buffer_rep.getvalue(),
            file_name=f"Reporte_{rango_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )                        

    except Exception as e:
        st.error(f"🚨 Fallo interno capturado: {e}")

if __name__ == "__main__":
    pass
