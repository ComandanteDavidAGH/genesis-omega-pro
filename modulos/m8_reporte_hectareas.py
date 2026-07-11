import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import io

# =================================================================
# 🛰️ BUCLE DE EXTRACCIÓN TOTAL DE LA OPERACIÓN
# =================================================================
def descargar_todo_supabase(_cliente_supabase):
    todos_los_datos = []
    inicio = 0
    paso = 1000
    
    while True:
        respuesta = _cliente_supabase.table("TABLA_1").select("*").range(inicio, inicio + paso - 1).execute()
        chunk = respuesta.data
        if not chunk:
            break
        todos_los_datos.extend(chunk)
        if len(chunk) < paso:
            break
        inicio += paso
        if inicio >= 40000: 
            break
            
    return todos_los_datos

# =================================================================
# 🚁 RADAR DE HECTÁREAS - OMEGA V19 (EDICIÓN DRON CORREGIDO)
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    
    st.markdown("<h1 style='color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: \"Arial Black\", sans-serif; text-transform: uppercase;'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
    col_emergencia, col_vacia = st.columns([2, 2])
    if col_emergencia.button("⚠️ LIMPIAR MEMORIA Y TRAER DATOS 2026", type="primary", use_container_width=True):
        if 'm8_datos_crudos' in st.session_state:
            del st.session_state['m8_datos_crudos']
        st.toast("Memoria RAM de la pestaña vaciada. Recargando Bóveda...", icon="🔄")
        st.rerun()

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

    def fmt_latino(val):
        try: return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except: return str(val) if val is not None else ""

    if supabase_client is None:
        st.error("🚨 Sin conexión a Supabase.")
        return

    if 'm8_datos_crudos' not in st.session_state:
        with st.spinner("🛰️ Extrayendo todo el historial Cloud..."):
            st.session_state['m8_datos_crudos'] = descargar_todo_supabase(supabase_client)

    raw_data = st.session_state['m8_datos_crudos']

    try:
        if not raw_data:
            st.warning("⚠️ No se encontraron registros en Supabase.")
            return

        datos_limpios = []
        for row in raw_data:
            r_norm = {str(k).replace("\n", " ").strip().upper(): (str(v).strip() if v is not None else "") for k, v in row.items()}
            
            llave_ha = next((k for k in r_norm.keys() if "FUMIG" in k), None)
            llave_hr = next((k for k in r_norm.keys() if "RENDIMIENTO (HORAS)" in k or "RENDIMIENTO  (HORAS)" in k), None)
            llave_sem = next((k for k in r_norm.keys() if k == "SEM" or k == "SEMANA"), None)
            
            f_dt = procesar_fecha_pesada(r_norm.get("FECHA", ""))
            if f_dt is None:
                continue

            datos_limpios.append({
                "PISTA": r_norm.get("PISTA", "").strip().upper(),
                "HK": r_norm.get("HK", "").strip().upper(),
                "MODELO": r_norm.get("MODELO", "").strip().upper(),
                "FECHA_REAL": f_dt,
                "SEMANA": r_norm.get(llave_sem, "") if llave_sem else "",
                "HA_NETAS": extraer_numero(r_norm.get(llave_ha, "0") if llave_ha else "0"),
                "H_PROPORCIONAL": extraer_numero(r_norm.get(llave_hr, "0") if llave_hr else "0")
            })

        df_rep = pd.DataFrame(datos_limpios)
        
        if df_rep.empty:
            st.warning("⚠️ Los datos de Supabase no tienen formatos de fecha válidos.")
            return

        pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
        
        # --- 🎛️ SELECTORES INDEPENDIENTES ---
        st.markdown("### 🎛️ Centro de Comando y Filtros")
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.4])
        
        vista_seleccionada = c1.radio("👁️ Vista:", ["📊 Resumen Gerencial", "📅 Mapa Semanal"], horizontal=True, key="m8_v_final_v3")
        
        fecha_sel_ini = c2.date_input("📅 Fecha Inicial:", value=date(2026, 1, 1), min_value=date(2024, 1, 1), max_value=date(2030, 12, 31), key="m8_dat_ini_v3")
        fecha_sel_fin = c3.date_input("📅 Fecha Final:", value=date(2026, 12, 31), min_value=date(2024, 1, 1), max_value=date(2030, 12, 31), key="m8_dat_fin_v3")
        
        pista_sel = c4.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp, key="m8_pista_v3")

        cc1, cc2, cc3 = st.columns(3)
        mostrar_horas = cc1.checkbox("⏱️ Mostrar Horas", value=True, key="m8_h_v3")
        calcular_rend_prom = cc2.checkbox("🚀 Mostrar Rend. (Ha/Hr)", value=True, key="m8_r_v3")
        agrupar_avion = cc3.toggle("✈️ Desglosar por Flota", value=False, key="m8_f_v3")

        st.info(f"📊 **Auditoría de Datos:** Registros cargados en memoria: **{len(df_rep)}** | Historial desde **{df_rep['FECHA_REAL'].min().strftime('%d/%m/%Y')}** hasta **{df_rep['FECHA_REAL'].max().strftime('%d/%m/%Y')}**")

        df_filt = df_rep[(df_rep['FECHA_REAL'] >= fecha_sel_ini) & (df_rep['FECHA_REAL'] <= fecha_sel_fin)].copy()
        if pista_sel != "TODAS":
            df_filt = df_filt[df_filt['PISTA'] == pista_sel]
        
        if df_filt.empty:
            st.warning(f"⚠️ No hay registros de vuelo en el rango del {fecha_sel_ini.strftime('%d/%m/%Y')} al {fecha_sel_fin.strftime('%d/%m/%Y')}.")
            return
            
        meses_nom = {1:"01-Ene", 2:"02-Feb", 3:"03-Mar", 4:"04-Abr", 5:"05-May", 6:"06-Jun", 7:"07-Jul", 8:"08-Ago", 9:"09-Sep", 10:"10-Oct", 11:"11-Nov", 12:"12-Dic"}
        df_filt['MES'] = df_filt['FECHA_REAL'].apply(lambda x: meses_nom.get(x.month, "Desconocido"))
        
        st.markdown("---")
        rango_txt = f"{fecha_sel_ini.strftime('%d/%m/%Y')} al {fecha_sel_fin.strftime('%d/%m/%Y')}"
        
        if vista_seleccionada == "📊 Resumen Gerencial":
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
                        
                        # 🎯 SE CORRIGE AQUÍ: "🛸" representa el Drone de forma tecnológica
                        emoji = "🛸 DRON:" if es_dron else "✈️ AVION:"
                        
                        fila_hk = {'NIVEL': '', 'AVIÓN (HK)': f"{emoji} {hk}", 'MES': 'Total Flota'}
                        if mostrar_horas or calcular_rend_prom: fila_hk['REND (hr)'] = sum_hr_hk
                        fila_hk['ÁREA FUMIG (ha)'] = sum_ha_hk
                        if calcular_rend_prom: fila_hk['PROMEDIO (Ha/Hr)'] = sum_hr_hk # Ajuste
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

            # --- 🎨 ESTILOS ---
            df_visual = pd.DataFrame(tabla_final)
            
            def aplicar_estilos_originales(row):
                if "BASE:" in str(row['NIVEL']):
                    return ['background-color: #d1ecf1; font-weight: bold; color: #0c5460;'] * len(row)
                elif "TOTAL GENERAL" in str(row['NIVEL']):
                    return ['background-color: #c3e6cb; font-weight: bold; color: #155724;'] * len(row)
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
            
            st.markdown("---")
            df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
            if not df_grafico.empty:
                df_grafico['TXT'] = df_grafico['TOTAL MES'].apply(fmt_latino)
                fig = px.bar(df_grafico, x='MES', y='TOTAL MES', text='TXT', color='TOTAL MES', color_continuous_scale='Greens')
                fig.update_traces(textposition='outside')
                fig.update_layout(xaxis_title="Mes", showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        # 🎯 EXPORTACIÓN EXCEL
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
            file_name=f"Reporte_Hectareas_{rango_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )                        

    except Exception as e:
        st.error(f"🚨 Fallo procesando el reporte: {e}")

if __name__ == "__main__":
    pass
