import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# =================================================================
# 🛡️ MEMORIA RAM BLINDADA (CACHÉ TÁCTICO)
# El guión bajo en _cliente_supabase evita que Streamlit colapse al leerlo
# =================================================================
@st.cache_data(ttl=600, show_spinner=False)
def descargar_datos_boveda(_cliente_supabase):
    respuesta = _cliente_supabase.table("TABLA_1").select("*").limit(15000).execute()
    return respuesta.data

# =================================================================
# 🚁 RADAR DE HECTÁREAS - OMEGA V12 FINAL
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
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
            try: return pd.to_datetime('1899-12-30') + pd.to_timedelta(int(texto), unit='D')
            except: pass
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y'):
            try: return datetime.strptime(texto, fmt)
            except: pass
        return None

    def fmt_latino(val):
        try: return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except: return str(val) if val is not None else ""

    if supabase_client is None:
        st.error("🚨 Sin conexión a Supabase.")
        return

    try:
        with st.spinner("🛰️ Accediendo a la Bóveda de Supabase (Memoria Caché Activada)..."):
            # 🎯 LLAMADA PROTEGIDA POR CACHÉ: Solo descarga 1 vez cada 10 minutos
            raw_data = descargar_datos_boveda(supabase_client)
            
        if not raw_data:
            st.warning("⚠️ La TABLA_1 está vacía en Supabase.")
            return

        # 🎯 PURIFICACIÓN EXTREMA: Sin saltos de línea ni NoneTypes
        datos_limpios = []
        for row in raw_data:
            r_norm = {str(k).replace("\n", " ").strip().upper(): (str(v).strip() if v is not None else "") for k, v in row.items()}
            
            llave_ha = next((k for k in r_norm.keys() if "FUMIG" in k), None)
            llave_hr = next((k for k in r_norm.keys() if "RENDIMIENTO (HORAS)" in k or "RENDIMIENTO  (HORAS)" in k), None)
            llave_sem = next((k for k in r_norm.keys() if k == "SEM" or k == "SEMANA"), None)
            
            datos_limpios.append({
                "OS": r_norm.get("Nº ORDEN", ""),
                "PISTA": r_norm.get("PISTA", "").upper(),
                "HK": r_norm.get("HK", "").upper(),
                "MODELO": r_norm.get("MODELO", "").upper(),
                "FECHA": r_norm.get("FECHA", ""),
                "SEMANA": r_norm.get(llave_sem, "") if llave_sem else "",
                "HA_NETAS": extraer_numero(r_norm.get(llave_ha, "0") if llave_ha else "0"),
                "H_PROPORCIONAL": extraer_numero(r_norm.get(llave_hr, "0") if llave_hr else "0")
            })

        df_rep = pd.DataFrame(datos_limpios)
        
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
            st.warning("⚠️ No se encontraron vuelos con Hectáreas y Fechas válidas. Revise la tabla.")
            return

        pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
        min_fecha_real = df_rep['FECHA_DT'].min().date()
        max_fecha_real = df_rep['FECHA_DT'].max().date()
        
        # --- 🎛️ PANEL DE CONTROL ---
        st.markdown("### 🎛️ Centro de Comando y Filtros")
        c1, c2, c3, c4 = st.columns([1.5, 1, 1, 1])
        
        vista_seleccionada = c1.radio("👁️ Vista:", ["📊 Resumen Gerencial", "📅 Mapa Semanal"], horizontal=True)
        fecha_sel_ini = c2.date_input("📅 Inicial:", value=min_fecha_real)
        fecha_sel_fin = c3.date_input("📅 Final:", value=max_fecha_real)
        pista_sel = c4.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp)
        
        mostrar_horas, calcular_rend_prom, agrupar_avion = False, False, False

        if vista_seleccionada == "📊 Resumen Gerencial":
            cc1, cc2, cc3 = st.columns(3)
            mostrar_horas = cc1.checkbox("⏱️ Mostrar Horas", value=True)
            calcular_rend_prom = cc2.checkbox("🚀 Mostrar Rend. (Ha/Hr)", value=True)
            agrupar_avion = cc3.toggle("✈️ Desglosar por Flota", value=False)

        df_filt = df_rep[(df_rep['FECHA_DT'].dt.date >= fecha_sel_ini) & (df_rep['FECHA_DT'].dt.date <= fecha_sel_fin)].copy()
        if pista_sel != "TODAS":
            df_filt = df_filt[df_filt['PISTA'] == pista_sel]
        
        if df_filt.empty:
            st.warning("⚠️ No hay datos en este rango.")
            return
            
        meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
        df_filt['MES'] = df_filt['FECHA_DT'].dt.month.map(meses_nom)
        
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
                        emoji = "🚁 DRON:" if es_dron else "✈️ AVION:"
                        
                        fila_hk = {'NIVEL': '', 'AVIÓN (HK)': f"{emoji} {hk}", 'MES': 'Total Flota'}
                        if mostrar_horas or calcular_rend_prom: fila_hk['REND (hr)'] = sum_hr_hk
                        fila_hk['ÁREA FUMIG (ha)'] = sum_ha_hk
                        if calcular_rend_prom: fila_hk['PROMEDIO (Ha/Hr)'] = sum_ha_hk / sum_hr_hk if sum_hr_hk > 0 else 0.0
                        tabla_final.append(fila_hk)
                        
                        for _, row in datos_hk.iterrows():
                            fila_mes = {'NIVEL': '', 'AVIÓN (HK)': '', 'MES': row['MES']}
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
                        fila_mes = {'NIVEL': '', 'MES': row['MES']}
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

            # TABLA NATIVA SEGURA
            df_visual = pd.DataFrame(tabla_final)
            for col in ['ÁREA FUMIG (ha)', 'REND (hr)', 'PROMEDIO (Ha/Hr)']:
                if col in df_visual.columns:
                    df_visual[col] = df_visual[col].apply(fmt_latino)
            
            st.dataframe(df_visual.astype(str), use_container_width=True, hide_index=True)

        else:
            matriz = pd.pivot_table(df_filt, values='HA_NETAS', index='MES', columns='SEMANA', aggfunc='sum', fill_value=0)
            matriz = matriz.sort_index()
            cols_ordenadas = sorted(matriz.columns, key=lambda x: int(x) if str(x).isdigit() else 999)
            matriz = matriz[cols_ordenadas]
            matriz.index = [m.split('-')[1] if '-' in m else m for m in matriz.index]
            matriz['TOTAL MES'] = matriz.sum(axis=1)
            matriz.loc['TOTAL ANUAL'] = matriz.sum(axis=0)
            
            st.markdown(f"#### 🛩️ Rendimiento Semana a Semana ({rango_txt})")
            
            df_matriz_str = matriz.copy()
            for col in df_matriz_str.columns:
                df_matriz_str[col] = df_matriz_str[col].apply(fmt_latino)
            st.dataframe(df_matriz_str.astype(str), use_container_width=True)
            
            st.markdown("---")
            df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
            if not df_grafico.empty:
                df_grafico['TXT'] = df_grafico['TOTAL MES'].apply(fmt_latino)
                fig = px.bar(df_grafico, x='index', y='TOTAL MES', text='TXT', color='TOTAL MES', color_continuous_scale='Greens')
                fig.update_traces(textposition='outside')
                fig.update_layout(xaxis_title="Mes", showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        # 🎯 EXPORTACIÓN EXCEL ULTRA-LIGERA
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
