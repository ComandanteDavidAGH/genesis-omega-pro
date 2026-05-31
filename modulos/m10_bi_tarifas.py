import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import gspread
import re
import math
import io
import openpyxl

# --- 🧪 APARTADO DE BARREDORAS Y AUXILIARES GLOBALES ---
def limpiar_encabezados(df):
    df.columns = [
        str(col).upper()
        .replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
        .replace('À','A').replace('È','E').replace('Ì','I').replace('Ò','O').replace('Ù','U')
        .strip()
        for col in df.columns
    ]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    if "" in df.columns: df = df.drop(columns=[""])
    return df
    
def estandarizar_base(df):
    renombres = {}
    for col in df.columns:
        col_u = str(col).upper().replace('\n', ' ').strip()
        if 'FACTURAR' in col_u:
            renombres[col] = 'COSTO_MAESTRO'
            break
            
    if 'COSTO_MAESTRO' not in renombres.values():
        for col in df.columns:
            col_u = str(col).upper().replace('\n', ' ').strip()
            if 'COSTO AVION ($/HA)' in col_u or col_u == 'COSTO_HA':
                renombres[col] = 'COSTO_MAESTRO'
                break
                
    finca_ok = False; fecha_ok = False; area_ok = False
    for col in df.columns:
        col_u = str(col).upper().replace('\n', ' ').strip()
        if not finca_ok and (col_u == 'FINCA' or col_u == 'PROPIEDAD'):
            renombres[col] = 'FINCA_MAESTRA'
            finca_ok = True
        elif not fecha_ok and col_u == 'FECHA':
            renombres[col] = 'FECHA_MAESTRA'
            fecha_ok = True
        elif not area_ok and ('FUMIG' in col_u or ('AREA' in col_u and 'BRUTA' not in col_u) or col_u == 'HAS'):
            renombres[col] = 'AREA_MAESTRA'
            area_ok = True
            
    df.rename(columns=renombres, inplace=True)
    return df
    
def convertir_pesos(val):
    try:
        v = str(val)
        v_limpio = "".join([c for c in v if c.isdigit() or c in ['.', ',']])
        v_limpio = v_limpio.rstrip('.,')
        if v_limpio == '': return 0.0
        
        if ',' in v_limpio and '.' not in v_limpio: v_limpio = v_limpio.replace(',', '.')
        elif '.' in v_limpio and ',' in v_limpio: v_limpio = v_limpio.replace('.', '').replace(',', '.')
        elif '.' in v_limpio:
            partes = v_limpio.split('.')
            if len(partes[-1]) == 3: v_limpio = v_limpio.replace('.', '')
                
        num = float(v_limpio)
        if 0 < num < 2000: num = num * 1000 
        return num
    except: return 0.0

def limpiar_area(val):
    try:
        v = str(val).upper().replace(',', '.')
        v = "".join([c for c in v if c.isdigit() or c == '.'])
        return float(v) if v != '' else 0.0
    except: return 0.0

def calcular_frecuencia(df):
    if df.empty or 'FECHA_DT' not in df.columns: return 0, 0
    fechas = sorted(df['FECHA_DT'].dt.date.unique())
    if not fechas: return 0, 0
    
    ciclos = 1
    inicios_ciclo = [fechas[0]]
    for i in range(1, len(fechas)):
        if (fechas[i] - fechas[i-1]).days > 5:
            ciclos += 1
            inicios_ciclo.append(fechas[i])
            
    if ciclos > 1:
        diffs = [(inicios_ciclo[j] - inicios_ciclo[j-1]).days for j in range(1, ciclos)]
        avg_int = sum(diffs) / len(diffs)
    else:
        avg_int = 0
    return ciclos, avg_int


# --- 📡 NÚCLEO OPERATIVO DEL DASHBOARD ESTRATÉGICO ---
def ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero):
    st.markdown("<h1 class='titulo-principal'>📊 Centro de Inteligencia Estratégica BI</h1>", unsafe_allow_html=True)
    st.markdown("### 🛰️ Panel de Auditoría y Comportamiento Histórico por Finca")
    st.info("🤖 **MOTOR IA BI:** Extrayendo memoria histórica y datos vivos...")

    try:
        with st.spinner("📡 Sincronizando Bóveda Maestra y Archivo Histórico (Motor Turbo)..."):
            url_act = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
            datos_brutos_act = descargar_matriz_rapida(url_act, "TABLA 1")
            
            if len(datos_brutos_act) > 5:
                df_vivos = pd.DataFrame(datos_brutos_act[5:], columns=datos_brutos_act[4])
                df_vivos = estandarizar_base(limpiar_encabezados(df_vivos))
                df_vivos['ORIGEN_BI'] = 'ACTUAL'
            else:
                df_vivos = pd.DataFrame()

            url_hist = "https://docs.google.com/spreadsheets/d/16OZdiWwW7nLHyZBEnhiKlDTDttR7Tjhn37O9zm6wJOk/edit"
            datos_brutos_hist = descargar_matriz_rapida(url_hist, "Datos")
            
            if len(datos_brutos_hist) > 0:
                df_historico = pd.DataFrame(datos_brutos_hist[1:], columns=datos_brutos_hist[0])
                df_historico = estandarizar_base(limpiar_encabezados(df_historico))
                df_historico['ORIGEN_BI'] = 'HISTORICO'
            else:
                df_historico = pd.DataFrame()

        if df_vivos.empty or df_historico.empty:
            st.warning("⚠️ Los sistemas de almacenamiento temporal están vacíos.")
            return

        columnas_comunes = list(set(df_vivos.columns).intersection(set(df_historico.columns)))
        if 'ORIGEN_BI' in columnas_comunes: columnas_comunes.remove('ORIGEN_BI')
        
        if 'COSTO_MAESTRO' not in columnas_comunes or 'FINCA_MAESTRA' not in columnas_comunes:
            st.error("🚨 Las columnas de indexación maestra fallaron. Verifique cabeceras en Drive.")
            return

        df_vivos_trim = df_vivos[columnas_comunes + ['ORIGEN_BI']].copy()
        df_historico_trim = df_historico[columnas_comunes + ['ORIGEN_BI']].copy()
        
        super_base_bi = pd.concat([df_historico_trim, df_vivos_trim], ignore_index=True)
        super_base_bi['FINCA_MAESTRA'] = super_base_bi['FINCA_MAESTRA'].astype(str).str.strip().str.upper()

        if 'FECHA_MAESTRA' not in super_base_bi.columns:
            st.error("🚨 No se ubicó un sensor cronológico válido.")
            return

        super_base_bi['FECHA_DT'] = super_base_bi['FECHA_MAESTRA'].apply(procesar_fecha_pesada)
        super_base_bi = super_base_bi.dropna(subset=['FECHA_DT'])
        
        super_base_bi['AÑO'] = super_base_bi['FECHA_DT'].dt.year.astype(int)
        super_base_bi['MES'] = super_base_bi['FECHA_DT'].dt.month.astype(int)
        super_base_bi['TRIMESTRE'] = super_base_bi['FECHA_DT'].dt.quarter.astype(int)
        
        fincas_disp = ["TODAS"] + sorted(super_base_bi['FINCA_MAESTRA'].dropna().unique().tolist())
        años_disp = sorted(super_base_bi['AÑO'].unique().tolist(), reverse=True)
        
        col_modelo = 'MODELO' if 'MODELO' in super_base_bi.columns else None
        if col_modelo:
            super_base_bi[col_modelo] = super_base_bi[col_modelo].astype(str).str.strip().str.upper()
            modelos_disp = ["TODOS"] + sorted(super_base_bi[col_modelo].unique().tolist())
        else:
            modelos_disp = ["TODOS"]
        
        f1, f2 = st.columns(2)
        finca_sel = f1.selectbox("📍 Objetivo Geográfico (Finca)", fincas_disp)
        modelo_sel = f2.selectbox("🚁 Escuadrón (Modelo/Tipo)", modelos_disp)
        
        t1, t2, t3, t4 = st.columns(4)
        idx_base = 1 if len(años_disp) > 1 else 0
        año_base = t1.selectbox("📅 Año Base (Referencia)", años_disp, index=idx_base)
        año_comp = t2.selectbox("📆 Año Actual (Evaluar)", años_disp, index=0)
        
        tipo_periodo = t3.selectbox("⏱️ Lupa Temporal", ["AÑO COMPLETO", "POR TRIMESTRE", "POR MES"])
        meses_dict = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
        
        if tipo_periodo == "POR TRIMESTRE":
            periodo_sel = t4.selectbox("📊 Seleccione Trimestre", [1, 2, 3, 4], format_func=lambda x: f"Q{x}")
            etiq_periodo = f"Q{periodo_sel}"
        elif tipo_periodo == "POR MES":
            periodo_sel = t4.selectbox("📅 Seleccione Mes", list(meses_dict.keys()), format_func=lambda x: meses_dict[x])
            etiq_periodo = meses_dict[periodo_sel]
        else:
            t4.markdown("<br><span style='color:gray;'>Visión Anual Activada</span>", unsafe_allow_html=True)
            periodo_sel = "TODOS"
            etiq_periodo = "Total"

        df_finca = super_base_bi.copy()
        if finca_sel != "TODAS": df_finca = df_finca[df_finca['FINCA_MAESTRA'] == finca_sel]
        if col_modelo and modelo_sel != "TODOS": df_finca = df_finca[df_finca[col_modelo] == modelo_sel]
            
        df_finca['COSTO_NUM'] = df_finca['COSTO_MAESTRO'].apply(convertir_pesos)

        col_avion_ha = None
        for col in df_finca.columns:
            col_u = str(col).upper().replace('Ó', 'O')
            if 'AVION' in col_u and ('/HA' in col_u or ' HA' in col_u or '(HA)' in col_u):
                col_avion_ha = col
                break
        
        if col_avion_ha:
            df_finca['AVION_NUM'] = df_finca[col_avion_ha].apply(convertir_pesos)
        else:
            df_finca['AVION_NUM'] = 0.0

        df_periodo_a = df_finca[df_finca['AÑO'] == año_base].copy()
        df_periodo_b = df_finca[df_finca['AÑO'] == año_comp].copy()
        
        if tipo_periodo == "POR TRIMESTRE":
            df_periodo_a = df_periodo_a[df_periodo_a['TRIMESTRE'] == periodo_sel]
            df_periodo_b = df_periodo_b[df_periodo_b['TRIMESTRE'] == periodo_sel]
        elif tipo_periodo == "POR MES":
            df_periodo_a = df_periodo_a[df_periodo_a['MES'] == periodo_sel]
            df_periodo_b = df_periodo_b[df_periodo_b['MES'] == periodo_sel]

        # =========================================================
        # 🎯 CORTAFUEGOS INTELIGENTE (LA CURA DEFINITIVA)
        # =========================================================
        col_area = 'AREA_MAESTRA' if 'AREA_MAESTRA' in df_finca.columns else None
        
        if col_area:
            df_periodo_a.loc[:, 'AREA_NUM'] = df_periodo_a[col_area].apply(limpiar_area)
            df_periodo_b.loc[:, 'AREA_NUM'] = df_periodo_b[col_area].apply(limpiar_area)
            
            # 1. Regresamos a la validación por FECHA y ÁREA para NO borrar vuelos legítimos
            df_vuelos_a = df_periodo_a.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM']).copy()
            df_vuelos_b = df_periodo_b.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM']).copy()
            
            # 2. El Cortafuegos Inteligente: Divide SOLO si el resultado tiene sentido
            def curar_costo(row):
                c = row['COSTO_NUM']
                a = row.get('AREA_NUM', 0)
                if pd.notna(c) and c > 350000 and pd.notna(a) and a > 0:
                    unit = c / a
                    # Si al dividirlo da un costo normal (entre 30k y 450k), entonces era facturación total
                    if 30000 <= unit <= 450000:
                        return unit
                return c

            df_vuelos_a['COSTO_NUM'] = df_vuelos_a.apply(curar_costo, axis=1)
            df_vuelos_b['COSTO_NUM'] = df_vuelos_b.apply(curar_costo, axis=1)
            
            df_periodo_a['COSTO_NUM'] = df_periodo_a.apply(curar_costo, axis=1)
            df_periodo_b['COSTO_NUM'] = df_periodo_b.apply(curar_costo, axis=1)

            area_a = df_vuelos_a['AREA_NUM'].sum() if not df_vuelos_a.empty else 0.0
            area_b = df_vuelos_b['AREA_NUM'].sum() if not df_vuelos_b.empty else 0.0
            
            costo_a = df_vuelos_a['COSTO_NUM'].mean() if not df_vuelos_a.empty else 0
            costo_b = df_vuelos_b['COSTO_NUM'].mean() if not df_vuelos_b.empty else 0
        else:
            # Respaldo si no hay columna de área
            df_vuelos_a = df_periodo_a.drop_duplicates(subset=['FECHA_DT'])
            df_vuelos_b = df_periodo_b.drop_duplicates(subset=['FECHA_DT'])
            area_a, area_b = 0.0, 0.0
            costo_a = df_vuelos_a['COSTO_NUM'].mean() if not df_vuelos_a.empty else 0
            costo_b = df_vuelos_b['COSTO_NUM'].mean() if not df_vuelos_b.empty else 0

        delta_pct = ((costo_b - costo_a) / costo_a * 100) if costo_a > 0 else 0
        
        st.markdown("### 📊 Auditoría de Costos: Impacto General por Hectárea")
        k1, k2, k3 = st.columns(3)
        k1.metric(label=f"Costo Promedio Ha ({año_base})", value=f"$ {costo_a:,.0f}")
        k2.metric(label=f"Costo Promedio Ha ({año_comp})", value=f"$ {costo_b:,.0f}")
        k3.metric(label="Variación Total (%)", value=f"{delta_pct:+.2f} %", delta=f"{delta_pct:+.2f}%", delta_color="inverse")
        
        st.markdown("#### 🚜 Volumen Operativo (Hectáreas Aplicadas)")
        var_area = ((area_b - area_a) / area_a * 100) if area_a > 0 else 0

        h1, h2, h3 = st.columns(3)
        h1.metric(f"Total Hectáreas ({año_base})", f"{area_a:,.1f} Ha")
        h2.metric(f"Total Hectáreas ({año_comp})", f"{area_b:,.1f} Ha")
        if area_a > 0: h3.metric("Variación de Área", f"{var_area:+.1f} %", delta=f"{var_area:+.1f}%", delta_color="normal")
        else: h3.metric("Variación de Área", "N/A")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if delta_pct > 10:
            st.error(f"⚠️ **ALERTA ROJA:** El costo operativo en {finca_sel} presenta una desviación del **{delta_pct:.1f}%**. Se requiere análisis de causa raíz.")
        elif delta_pct < 0:
            st.success(f"✅ **RENDIMIENTO ÓPTIMO:** El costo operativo se redujo. Excelente gestión logística.")
        else:
            st.info(f"⚖️ **ESTABILIDAD:** Los costos se mantienen dentro de los márgenes normales de variación.")
            
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### ⏱️ Análisis de Frecuencia: Ciclos Reales e Intervalo")
        
        ciclos_a, int_a = calcular_frecuencia(df_periodo_a)
        ciclos_b, int_b = calcular_frecuencia(df_periodo_b)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(f"Ciclos ({año_base})", f"{ciclos_a} ciclos")
        c2.metric(f"Ciclos ({año_comp})", f"{ciclos_b} ciclos", delta=f"{ciclos_b - ciclos_a} ciclos", delta_color="inverse")
        str_int_a = f"{int_a:.1f} días" if int_a > 0 else "N/A"
        str_int_b = f"{int_b:.1f} días" if int_b > 0 else "N/A"
        c3.metric(f"Intervalo Prom. ({año_base})", str_int_a)
        if int_a > 0 and int_b > 0:
            delta_int = int_b - int_a
            c4.metric(f"Intervalo Prom. ({año_comp})", str_int_b, delta=f"{delta_int:+.1f} días", delta_color="normal")
        else:
            c4.metric(f"Intervalo Prom. ({año_comp})", str_int_b)
        
        st.markdown("---")
        st.markdown("### 🧬 Análisis de Causa Raíz: Atribución de Variaciones")
        
        df_tendencia = pd.concat([df_periodo_a, df_periodo_b])
        if not df_tendencia.empty:
            if col_area:
                df_tend_unicos = df_tendencia.drop_duplicates(subset=['FECHA_DT', 'AREA_NUM'])
            else:
                df_tend_unicos = df_tendencia.drop_duplicates(subset=['FECHA_DT'])

            if tipo_periodo in ["AÑO COMPLETO", "POR TRIMESTRE"]:
                tendencia_agrupa = df_tend_unicos.groupby(['AÑO', 'MES'])['COSTO_NUM'].mean().reset_index()
                tendencia_agrupa['EJE_X'] = tendencia_agrupa['MES'].map(meses_dict)
                tendencia_agrupa = tendencia_agrupa.sort_values('MES')
                titulo_x = "Meses Operativos"
            else:
                df_tend_unicos['DIA'] = df_tend_unicos['FECHA_DT'].dt.day
                tendencia_agrupa = df_tend_unicos.groupby(['AÑO', 'DIA'])['COSTO_NUM'].mean().reset_index()
                tendencia_agrupa['EJE_X'] = "Día " + tendencia_agrupa['DIA'].astype(str)
                tendencia_agrupa = tendencia_agrupa.sort_values('DIA')
                titulo_x = f"Días Operativos ({etiq_periodo})"
                
            tendencia_agrupa['AÑO'] = tendencia_agrupa['AÑO'].astype(str)
            fig_tendencia = px.line(tendencia_agrupa, x='EJE_X', y='COSTO_NUM', color='AÑO', markers=True, color_discrete_sequence=['#2F75B5', '#ef4444'])
            fig_tendencia.update_layout(yaxis_title="Costo Promedio ($ COP / Ha)", xaxis_title=titulo_x, plot_bgcolor='rgba(0,0,0,0)', hovermode="x unified")
            max_y = tendencia_agrupa['COSTO_NUM'].max() * 1.2
            if not pd.isna(max_y): fig_tendencia.update_yaxes(range=[0, max_y])
            fig_tendencia.update_traces(line=dict(width=3), marker=dict(size=8), texttemplate="$ %{y:,.0f}", textposition="top center", hovertemplate="<b>%{x}</b><br>Costo: $ %{y:,.0f} COP/Ha<extra></extra>")
            st.plotly_chart(fig_tendencia, use_container_width=True)
        else:
            st.warning("⚠️ No hay suficientes operaciones en este periodo exacto para trazar una curva comparativa.")
            
        st.markdown("<hr>", unsafe_allow_html=True)
        
        vuelo_a = df_vuelos_a['AVION_NUM'].mean() if not df_vuelos_a.empty else 0
        vuelo_b = df_vuelos_b['AVION_NUM'].mean() if not df_vuelos_b.empty else 0
        
        insumos_a = max(0, costo_a - vuelo_a)
        insumos_b = max(0, costo_b - vuelo_b)

        vuelo_tot_a = vuelo_a * area_a
        vuelo_tot_b = vuelo_b * area_b
        insumos_tot_a = insumos_a * area_a
        insumos_tot_b = insumos_b * area_b

        st.markdown("#### 🛩️ vs 🧪 Distribución del Encarecimiento")
        categorias = [f'Análisis {año_base}', f'Análisis {año_comp}']
        tab_unit, tab_glob = st.tabs(["🎯 Impacto Unitario (Promedio / Ha)", "💰 Impacto Global (Presupuesto Total)"])
        
        with tab_unit:
            fig_unit = go.Figure(data=[
                go.Bar(name='Costo Avión / Ha', x=categorias, y=[vuelo_a, vuelo_b], marker_color='#2F75B5', text=[f"$ {vuelo_a:,.0f}", f"$ {vuelo_b:,.0f}"], textposition='auto'),
                go.Bar(name='Costo Insumos / Ha', x=categorias, y=[insumos_a, insumos_b], marker_color='#548235', text=[f"$ {insumos_a:,.0f}", f"$ {insumos_b:,.0f}"], textposition='auto')
            ])
            fig_unit.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor COP / Ha", margin=dict(t=20, b=20))
            st.plotly_chart(fig_unit, use_container_width=True)
            
        with tab_glob:
            fig_glob = go.Figure(data=[
                go.Bar(name='Total Facturación Avión', x=categorias, y=[vuelo_tot_a, vuelo_tot_b], marker_color='#2F75B5', text=[f"$ {vuelo_tot_a:,.0f}", f"$ {vuelo_tot_b:,.0f}"], textposition='auto'),
                go.Bar(name='Total Consumo Insumos', x=categorias, y=[insumos_tot_a, insumos_tot_b], marker_color='#548235', text=[f"$ {insumos_tot_a:,.0f}", f"$ {insumos_tot_b:,.0f}"], textposition='auto')
            ])
            fig_glob.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', yaxis_title="Valor Total COP", margin=dict(t=20, b=20))
            st.plotly_chart(fig_glob, use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Desglose Operativo: Cócteles y Variación")
        col_coctel = 'COCTEL' if 'COCTEL' in df_finca.columns else ('COCTEL_MAESTRO' if 'COCTEL_MAESTRO' in df_finca.columns else None)
        col_gln = 'GLN_HA' if 'GLN_HA' in df_finca.columns else None
        
        if col_coctel:
            df_periodo_a.loc[:, col_coctel] = df_periodo_a[col_coctel].astype(str).str.strip().str.upper()
            df_periodo_b.loc[:, col_coctel] = df_periodo_b[col_coctel].astype(str).str.strip().str.upper()
            
            agg_dict = {'COSTO_NUM': 'mean'}
            if col_gln: agg_dict[col_gln] = 'mean'
            
            g_a = df_periodo_a.groupby(col_coctel).agg(agg_dict).reset_index()
            g_b = df_periodo_b.groupby(col_coctel).agg(agg_dict).reset_index()
            
            tabla_autopsia = pd.merge(g_a, g_b, on=col_coctel, how='outer', suffixes=('_BASE', '_ACTUAL'))
            tabla_autopsia.fillna(0, inplace=True)
            
            tabla_autopsia.rename(columns={col_coctel: 'CÓCTEL APLICADO', 'COSTO_NUM_BASE': f'Costo/Ha ({año_base})', 'COSTO_NUM_ACTUAL': f'Costo/Ha ({año_comp})'}, inplace=True)
            tabla_autopsia['Variación ($)'] = tabla_autopsia[f'Costo/Ha ({año_comp})'] - tabla_autopsia[f'Costo/Ha ({año_base})']
            
            if col_gln:
                tabla_autopsia.rename(columns={f'{col_gln}_BASE': f'Gln/Ha ({año_base})', f'{col_gln}_ACTUAL': f'Gln/Ha ({año_comp})'}, inplace=True)
                
            df_vista = tabla_autopsia.copy()
            df_vista[f'Costo/Ha ({año_base})'] = df_vista[f'Costo/Ha ({año_base})'].map("$ {:,.0f}".format)
            df_vista[f'Costo/Ha ({año_comp})'] = df_vista[f'Costo/Ha ({año_comp})'].map("$ {:,.0f}".format)
            df_vista['Variación ($)'] = df_vista['Variación ($)'].map("$ {:,.0f}".format)
            st.dataframe(df_vista, use_container_width=True)
            
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### 🔬 Nivel 2: Composición del Cóctel y Variación Real de Insumos")

        if col_coctel:
            cocteles_disponibles = sorted(list(set(df_periodo_a[col_coctel].dropna().unique()) | set(df_periodo_b[col_coctel].dropna().unique())))
            coctel_sel = st.selectbox("🎯 Seleccione un Cóctel para auditar su receta año vs año:", ["SELECCIONE UN CÓCTEL..."] + cocteles_disponibles)

            if coctel_sel != "SELECCIONE UN CÓCTEL...":
                with st.spinner("Desplegando Deliberador IA..."):
                    if "gcp_credentials" in st.secrets: gc_rec = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else: gc_rec = gspread.service_account(filename='credenciales.json')
                    
                    df_mezclas = pd.DataFrame()
                    boveda_recetas = gc_rec.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    hoja_mezclas = boveda_recetas.worksheet("DD_Mesclas")
                    data_mez = hoja_mezclas.get('A:D')
                    if data_mez and len(data_mez) > 1:
                        df_mezclas = pd.DataFrame(data_mez[1:], columns=data_mez[0])
                        df_mezclas['COCTEL_CLEAN'] = df_mezclas.iloc[:,0].astype(str).str.upper().str.replace(" ", "")
                    
                    df_conf = pd.DataFrame(boveda_recetas.worksheet("Configuración").get_all_values()[1:], columns=boveda_recetas.worksheet("Configuración").get_all_values()[0])
                    df_dicc = pd.DataFrame(boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()[1:], columns=boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()[0])

                    url_precios = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
                    sh_precios = gc_rec.open_by_url(url_precios)
                    
                    precios_consolidados = []
                    for ws in sh_precios.worksheets():
                        datos_hoja = ws.get_all_values()
                        if not datos_hoja: continue
                        idx_header, col_anio, col_prod = -1, -1, -1
                        for i in range(min(10, len(datos_hoja))):
                            fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                            if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                                idx_header = i; col_anio = fila_upper.index('AÑO'); col_prod = fila_upper.index('PRODUCTO')
                                break
                        if idx_header != -1:
                            for row in datos_hoja[idx_header+1:]:
                                if len(row) > max(col_anio, col_prod):
                                    anio_str = str(row[col_anio]).strip()
                                    prod_str = str(row[col_prod]).strip().upper()
                                    if anio_str and prod_str:
                                        col_inicio_semanas = max(col_anio, col_prod) + 1
                                        valores_semana = [extraer_numero(v) for v in row[col_inicio_semanas:] if extraer_numero(v) > 0]
                                        promedio = sum(valores_semana)/len(valores_semana) if valores_semana else 0.0
                                        precios_consolidados.append({'AÑO': anio_str, 'PRODUCTO': prod_str, 'PRECIO_PROM': promedio})

                    df_precios = pd.DataFrame(precios_consolidados)
                    coctel_crudo = coctel_sel.upper().replace(" ", "")
                    partes_coctel = coctel_crudo.split('+')
                    base_coctel = partes_coctel[0]
                    aditivos = partes_coctel[1:] if len(partes_coctel) > 1 else []

                    match_num = re.search(r'\d+', base_coctel)
                    dosis_aceite = int(match_num.group()) if match_num else 0
                    solo_letras = re.sub(r'\d+', '', base_coctel)

                    dict_prods_unicos = {}
                    es_organico = False

                    try:
                        df_t2 = pd.DataFrame(boveda_recetas.worksheet("TABLA 2").get_all_values()[1:], columns=boveda_recetas.worksheet("TABLA 2").get_all_values()[0])
                        match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_sel.upper().strip()]
                        if not match_f.empty and "ORGANIC" in str(match_f.iloc[0, 5]).upper(): es_organico = True
                    except: pass

                    receta_base = pd.DataFrame()
                    if not df_mezclas.empty:
                        if es_organico and not base_coctel.endswith('O'):
                            coctel_prueba = f"{base_coctel}O"
                            if not df_mezclas[df_mezclas['COCTEL_CLEAN'] == coctel_prueba].empty: base_coctel = coctel_prueba
                        receta_base = df_mezclas[df_mezclas['COCTEL_CLEAN'] == base_coctel]
                        if receta_base.empty: receta_base = df_mezclas[df_mezclas['COCTEL_CLEAN'] == solo_letras]

                    if not receta_base.empty:
                        for idx, row in receta_base.iterrows():
                            prod = str(row.iloc[1]).strip().upper()
                            dosis = extraer_numero(row.iloc[2])
                            if dosis > 0 and prod not in ['NAN', '']: dict_prods_unicos[prod] = dosis
                    else:
                        if not df_dicc.empty:
                            siglas_validas = df_dicc[df_dicc['SIGLA'].astype(str).str.strip() != '']['SIGLA'].astype(str).str.strip().str.upper().unique().tolist()
                            siglas_validas.sort(key=len, reverse=True)
                            resto_letras = solo_letras
                            for sigla in siglas_validas:
                                if sigla in resto_letras:
                                    match_sig = df_dicc[df_dicc['SIGLA'].astype(str).str.strip().str.upper() == sigla]
                                    if not match_sig.empty:
                                        prod_name = str(match_sig.iloc[0]['PRODUCTO']).strip().upper()
                                        dict_prods_unicos[prod_name] = extraer_numero(match_sig.iloc[0]['DOSIS'])
                                        resto_letras = resto_letras.replace(sigla, '', 1)
                            if dosis_aceite > 0: dict_prods_unicos['ACEITE DICAM'] = float(dosis_aceite)
                            dict_prods_unicos['ACONDICIONADOR SV'] = 0.02
                            dict_prods_unicos['ADHERENTE SV'] = 0.13

                    if not df_dicc.empty:
                        for ad in aditivos:
                            match_sig = df_dicc[df_dicc['SIGLA'].astype(str).str.strip().str.upper() == ad]
                            if not match_sig.empty:
                                prod_name = str(match_sig.iloc[0]['PRODUCTO']).strip().upper()
                                dict_prods_unicos[prod_name] = extraer_numero(match_sig.iloc[0]['DOSIS'])
                            else:
                                if "ZN" in ad: dict_prods_unicos["ZINTRAC"] = 0.5
                                elif "BT" in ad: dict_prods_unicos["BANATREL"] = 0.5

                    if not df_dicc.empty:
                        for prod_name in list(dict_prods_unicos.keys()):
                            match_dicc = df_dicc[df_dicc['PRODUCTO'].astype(str).str.strip().str.upper() == prod_name]
                            if not match_dicc.empty:
                                if 'ORGANIC' in str(match_dicc.iloc[0].get('TIPO DE CULTIVO', '')).upper(): es_organico = True

                    if dosis_aceite > 0:
                        aceite_key = next((k for k in dict_prods_unicos.keys() if "ACEITE" in k), "ACEITE DICAM")
                        dict_prods_unicos[aceite_key] = float(dosis_aceite)
                    else:
                        for k in [k for k in dict_prods_unicos.keys() if "ACEITE" in k]: dict_prods_unicos.pop(k, None)

                    if es_organico:
                        for k in [k for k in dict_prods_unicos.keys() if "ADHERENTE" in k]: dict_prods_unicos.pop(k, None)
                        sprayfix_key = next((k for k in dict_prods_unicos.keys() if "SPRAYFIX" in k), "SPRAYFIX")
                        if sprayfix_key not in dict_prods_unicos: dict_prods_unicos[sprayfix_key] = 0.2
                    else:
                        for k in [k for k in dict_prods_unicos.keys() if "SPRAYFIX" in k]: dict_prods_unicos.pop(k, None)
                        adherente_key = next((k for k in dict_prods_unicos.keys() if "ADHERENTE" in k), "ADHERENTE SV")
                        if adherente_key not in dict_prods_unicos: dict_prods_unicos[adherente_key] = 0.13

                    prods_receta = [{"PRODUCTO": k, "DOSIS": v} for k, v in dict_prods_unicos.items() if v > 0]
                    if prods_receta:
                        matriz_mol = []
                        
                        def obtener_precio_promedio(producto, anio_obj):
                            if not df_precios.empty:
                                match_df = df_precios[(df_precios['AÑO'] == str(anio_obj)) & (df_precios['PRODUCTO'] == producto)]
                                if match_df.empty: match_df = df_precios[(df_precios['AÑO'] == str(anio_obj)) & (df_precios['PRODUCTO'].str.contains(producto))]
                                if not match_df.empty and match_df['PRECIO_PROM'].mean() > 0: return match_df['PRECIO_PROM'].mean()
                            if str(anio_obj) == str(año_comp) or str(anio_obj) == str(datetime.now().year):
                                match_conf = df_conf[df_conf.iloc[:, 8].astype(str).str.upper().str.strip() == producto]
                                if match_conf.empty: match_conf = df_conf[df_conf.iloc[:, 8].astype(str).str.upper().str.strip().str.contains(producto)]
                                if not match_conf.empty: return extraer_numero(match_conf.iloc[0, 9])
                            return 0.0

                        costo_total_a, costo_total_b = 0.0, 0.0
                        for item in prods_receta:
                            prod, dosis = item["PRODUCTO"], item["DOSIS"]
                            precio_a, precio_b = obtener_precio_promedio(prod, año_base), obtener_precio_promedio(prod, año_comp)
                            costo_ha_a, costo_ha_b = dosis * precio_a, dosis * precio_b
                            costo_total_a += costo_ha_a
                            costo_total_b += costo_ha_b
                            matriz_mol.append({"INSUMO QUÍMICO": prod, "DOSIS/HA": f"{dosis:.3f}", f"P. Prom. ({año_base})": f"$ {precio_a:,.0f}", f"P. Prom. ({año_comp})": f"$ {precio_b:,.0f}", f"Costo/Ha ({año_base})": costo_ha_a, f"Costo/Ha ({año_comp})": costo_ha_b, "Variación ($)": costo_ha_b - costo_ha_a})

                        df_vista_mol = pd.DataFrame(matriz_mol).sort_values('Variación ($)', ascending=False)
                        df_vista_mol[f"Costo/Ha ({año_base})"] = df_vista_mol[f"Costo/Ha ({año_base})"].map("$ {:,.0f}".format)
                        df_vista_mol[f"Costo/Ha ({año_comp})"] = df_vista_mol[f"Costo/Ha ({año_comp})"].map("$ {:,.0f}".format)
                        df_vista_mol["Variación ($)"] = df_vista_mol["Variación ($)"].map("$ {:,.0f}".format)
                        st.dataframe(df_vista_mol, use_container_width=True, hide_index=True)
                        
                        c1, c2, c3 = st.columns(3)
                        c1.metric(f"Total Teórico ({año_base})", f"$ {costo_total_a:,.0f}")
                        c2.metric(f"Total Teórico ({año_comp})", f"$ {costo_total_b:,.0f}")
                        c3.metric("Variación Cóctel", f"$ {costo_total_b - costo_total_a:,.0f}", delta=f"$ {costo_total_b - costo_total_a:,.0f}", delta_color="inverse")
                        
                        if 'AVION_NUM' in df_periodo_b.columns:
                            df_coctel_b = df_periodo_b[df_periodo_b[col_coctel] == coctel_sel]
                            costo_total_facturado_b = df_coctel_b['COSTO_NUM'].mean() if not df_coctel_b.empty else 0
                            vuelo_facturado_b = df_coctel_b['AVION_NUM'].mean() if not df_coctel_b.empty else 0
                            insumos_facturados_b = max(0, costo_total_facturado_b - vuelo_facturado_b)
                            
                            if costo_total_b > 0 and insumos_facturados_b > 0:
                                diff_b = insumos_facturados_b - costo_total_b
                                st.markdown("---")
                                st.markdown("### 🤖 Deliberador IA: Auditoría de Facturación SAP vs Receta Teórica")
                                if abs(diff_b) <= 2000: st.success(f"✅ **AUDITORÍA PERFECTA:** El costo de químicos facturados en SAP ($ {insumos_facturados_b:,.0f}) coincide con la receta ($ {costo_total_b:,.0f}).")
                                else:
                                    st.warning(f"⚠️ **DISCREPANCIA DETECTADA:** Los insumos facturados ($ {insumos_facturados_b:,.0f}) no cuadran con el teórico ($ {costo_total_b:,.0f}). Diferencia: **$ {diff_b:,.0f} / Ha**")
                                    st.markdown("#### 🔍 Conclusiones del Deliberador:")
                                    if diff_b > 0: st.write(f"- 📈 **Sobrecosto:** Se cobró más de lo que indica la sigla. Es probable que se haya aplicado **SPRAYFIX**, **ADHERENTE** extra o mayor dosis de **ACEITE**.")
                                    else: st.write(f"- 📉 **Ahorro/Faltante:** Se cobró menos. Si la finca es orgánica, se facturó correctamente (sin adherente), o hubo un error a favor en SAP.")
                                    
                                    if not df_precios.empty:
                                        st.write("- **Posibles causantes de la diferencia:**")
                                        candidatos_encontrados = False
                                        for idx, p_row in df_precios[df_precios['AÑO'] == str(año_comp)].iterrows():
                                            precio_p = p_row['PRECIO_PROM']
                                            for d in [0.02, 0.06, 0.13, 0.2, 0.5, 1.0, 2.0]:
                                                costo_teorico = precio_p * d
                                                if costo_teorico > 0 and abs(costo_teorico - abs(diff_b)) <= (abs(diff_b) * 0.15 + 500):
                                                    st.info(f"💡 ¿Se aplicó/omitió **{p_row['PRODUCTO']}** a dosis de **{d} L/Ha**? (Costo aprox: $ {costo_teorico:,.0f})")
                                                    candidatos_encontrados = True; break
                                            if candidatos_encontrados: break
                                        if not candidatos_encontrados: st.write("No se detectó un químico individual que coincida exacto.")
                    else: st.info("No se encontraron ingredientes válidos para esta receta.")
            else: st.warning("⚠️ No se encontró la columna 'COCTEL' en la base fusionada.")
                
        # =====================================================================
        # --- 🤝 NUEVO: SIMULADOR DE NEGOCIACIÓN Y AUDITORÍA DE TARIFAS ---
        # =====================================================================
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### 🤝 Simulador de Negociación (Tarifas de Aerofumigación)")
        st.info("💡 RADAR BLINDADO: Extracción estricta de Tarifas Unitarias (Avión + Dominical) asegurada por Fecha + Finca.")

        with st.container():
            c_sim1, c_sim2, c_sim3 = st.columns(3)
            col_pista_sim = next((c for c in super_base_bi.columns if "PISTA" in c or "ALMACEN" in c), None)
            if col_pista_sim:
                pistas_sim_disp = ["TODAS"] + sorted(super_base_bi[col_pista_sim].dropna().astype(str).str.upper().unique().tolist())
            else:
                pistas_sim_disp = ["TODAS"]

            sim_fecha_inicio = c_sim1.date_input("📅 Fecha Inicial:", value=datetime.now().date(), key="sim_f_ini_f")
            sim_fecha_fin = c_sim2.date_input("📅 Fecha Final:", value=datetime.now().date(), key="sim_f_fin_f")
            sim_pista = c_sim3.selectbox("📍 Base / Pista:", pistas_sim_disp, key="sim_pista_f")

        st.markdown("<br>", unsafe_allow_html=True)
        c_sim_m1, c_sim_m2, c_sim_m3 = st.columns(3)
        margen_actual = c_sim_m1.number_input("📉 Margen Actual en Factura (%)", value=8.0, step=0.5, key="marg_act_f")
        margen_nuevo = c_sim_m2.number_input("📈 Nuevo Margen a Simular (%)", value=11.0, step=0.5, key="marg_nue_f")
        
        with c_sim_m3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_simular = st.button("🚀 EJECUTAR SIMULACIÓN", type="primary", use_container_width=True, key="btn_simular_f")

        if btn_simular:
            with st.spinner("Procesando auditoría de simulación..."):
                df_sim = super_base_bi.copy()
                df_sim = df_sim[(df_sim['FECHA_DT'].dt.date >= sim_fecha_inicio) & (df_sim['FECHA_DT'].dt.date <= sim_fecha_fin)]
                if col_pista_sim and sim_pista != "TODAS":
                    df_sim = df_sim[df_sim[col_pista_sim].astype(str).str.upper() == sim_pista]

                col_ha = 'AREA_MAESTRA'
                if col_ha in df_sim.columns:
                    df_sim[col_ha] = pd.to_numeric(df_sim[col_ha].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
                    df_sim = df_sim[df_sim[col_ha] > 0]

                if df_sim.empty:
                    st.warning("⚠️ No se encontraron Órdenes de Servicio para el rango de fechas seleccionado.")
                else:
                    def red_excel(num):
                        return math.floor(num + 0.5) if num >= 0 else math.ceil(num - 0.5)

                    col_tarifa_avion, col_dominical = None, None
                    for c in df_sim.columns:
                        c_upper = str(c).upper().strip()
                        if "AVION" in c_upper and "HA" in c_upper and "COSTO" in c_upper: col_tarifa_avion = c
                        if "DOMINIC" in c_upper: col_dominical = c
                        
                    if not col_tarifa_avion:
                        for c in df_sim.columns:
                            c_upper = str(c).upper().strip()
                            if "AVION" in c_upper and "HA" in c_upper: col_tarifa_avion = c; break

                    col_finca = 'FINCA_MAESTRA'
                    
                    df_sim_unicos = df_sim.drop_duplicates(subset=['FECHA_DT', col_finca, col_ha, col_tarifa_avion])

                    matriz_simulacion = []

                    for _, row in df_sim_unicos.iterrows():
                        finca_val = str(row[col_finca]).upper().strip()
                        ha_val = float(row[col_ha])
                        pista_val = str(row[col_pista_sim]).upper().strip() if col_pista_sim else "N/A"
                        
                        col_os = next((c for c in df_sim.columns if "OS" in str(c).upper() and "COSTO" not in str(c).upper()), df_sim.columns[0])
                        for c in df_sim.columns:
                            if str(c).upper().strip() in ["OS", "ORDEN", "Nº OS", "Nº ORDEN", "ORDEN DE SERVICIO"]:
                                col_os = c; break
                        os_val = str(row[col_os]).strip()
                        
                        if pd.notna(row['FECHA_DT']):
                            fecha_val = row['FECHA_DT'].strftime('%d/%m/%Y')
                            semana_val = (row['FECHA_DT'] + pd.Timedelta(days=2)).isocalendar()[1]
                        else:
                            fecha_val = str(row['FECHA_MAESTRA'])
                            col_sem = next((c for c in df_sim.columns if "SEMANA" in str(c).upper()), None)
                            semana_val = row[col_sem] if col_sem else "N/A"

                        tar_avion_raw = convertir_pesos(row[col_tarifa_avion]) if col_tarifa_avion and col_tarifa_avion in row else 0.0
                        tar_dom_raw = convertir_pesos(row[col_dominical]) if col_dominical and col_dominical in row else 0.0
                        
                        tarifa_unitaria_actual = tar_avion_raw + tar_dom_raw

                        if tarifa_unitaria_actual > 0 and ha_val > 0:
                            t_act_red = red_excel(tarifa_unitaria_actual)
                            base_neta_ha = tarifa_unitaria_actual / (1 + (margen_actual / 100))
                            tarifa_nueva_unitaria = base_neta_ha * (1 + (margen_nuevo / 100))
                            t_nue_red = red_excel(tarifa_nueva_unitaria)
                            resta_tarifas = t_nue_red - t_act_red
                            
                            diferencia_total = red_excel(resta_tarifas * ha_val)
                            total_actual = red_excel(t_act_red * ha_val)
                            total_nuevo = red_excel(t_nue_red * ha_val)

                            matriz_simulacion.append({
                                "Nº OS": os_val, 
                                "FECHA": fecha_val, 
                                "SEMANA": int(semana_val) if str(semana_val).isdigit() else semana_val, 
                                "FINCA": finca_val, 
                                "PISTA": pista_val, 
                                "HECTÁREAS": ha_val, 
                                f"TARIFA ACTUAL / Ha ({margen_actual}%)": t_act_red, 
                                f"NUEVA TARIFA / Ha ({margen_nuevo}%)": t_nue_red, 
                                "TOTAL ACTUAL ($)": total_actual, 
                                "NUEVO TOTAL ($)": total_nuevo, 
                                "DIFERENCIA ($)": diferencia_total
                            })

                    if not matriz_simulacion:
                        st.warning("⚠️ No se pudieron extraer tarifas válidas. Verifique columnas T y U.")
                    else:
                        df_resultados = pd.DataFrame(matriz_simulacion)
                        df_semanal = df_resultados.groupby("SEMANA").agg({"HECTÁREAS": "sum", "TOTAL ACTUAL ($)": "sum", "NUEVO TOTAL ($)": "sum", "DIFERENCIA ($)": "sum"}).reset_index().sort_values(by="SEMANA").reset_index(drop=True)

                        total_actual_global = df_resultados["TOTAL ACTUAL ($)"].sum()
                        total_simulado_global = df_resultados["NUEVO TOTAL ($)"].sum()
                        total_diferencia_global = df_resultados["DIFERENCIA ($)"].sum()

                        st.markdown("### 🎯 Impacto Financiero Real de la Simulación")
                        k1, k2, k3 = st.columns(3)
                        k1.metric(f"💰 Total Actual ({margen_actual}%)", f"$ {total_actual_global:,.0f}")
                        k2.metric(f"📈 Proyección ({margen_nuevo}%)", f"$ {total_simulado_global:,.0f}")
                        k3.metric("⚖️ Dinero Real en Juego", f"$ {abs(total_diferencia_global):,.0f}", delta=f"$ {total_diferencia_global:,.0f}", delta_color="normal" if total_diferencia_global > 0 else "inverse")

                        st.markdown("<br>", unsafe_allow_html=True)
                        tab_resumen, tab_detalle = st.tabs(["📊 1. Resumen Macroeconómico", "📋 2. Desglose Quirúrgico"])
                        
                        with tab_resumen:
                            df_sem_vista = df_semanal.copy()
                            df_sem_vista["HECTÁREAS"] = df_sem_vista["HECTÁREAS"].map("{:,.2f}".format)
                            for col in ["TOTAL ACTUAL ($)", "NUEVO TOTAL ($)", "DIFERENCIA ($)"]: df_sem_vista[col] = df_sem_vista[col].map("$ {:,.0f}".format)
                            st.dataframe(df_sem_vista, use_container_width=True, hide_index=True)
                        
                        with tab_detalle:
                            df_vista = df_resultados.copy()
                            df_vista["HECTÁREAS"] = df_vista["HECTÁREAS"].map("{:,.2f}".format)
                            for col in [f"TARIFA ACTUAL / Ha ({margen_actual}%)", f"NUEVA TARIFA / Ha ({margen_nuevo}%)", "TOTAL ACTUAL ($)", "NUEVO TOTAL ($)", "DIFERENCIA ($)"]: df_vista[col] = df_vista[col].map("$ {:,.0f}".format)
                            
                            def col_dif(val):
                                if isinstance(val, str) and "-" in val: return 'color: #721c24; background-color: #f8d7da; font-weight: bold; text-align: center;'
                                if isinstance(val, str) and "$" in val: return 'color: #155724; background-color: #d4edda; font-weight: bold; text-align: center;'
                                return ''
                            st.dataframe(df_vista.style.map(col_dif, subset=["DIFERENCIA ($)"]), use_container_width=True, hide_index=True)

                        buffer_neg = io.BytesIO()
                        with pd.ExcelWriter(buffer_neg, engine='openpyxl') as writer:
                            df_semanal.to_excel(writer, sheet_name='Resumen_Semanal', index=False)
                            df_resultados.to_excel(writer, sheet_name='Detalle_OS', index=False)
                            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                            borde = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                            fondo, blanca = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid"), Font(color="FFFFFF", bold=True)
                            for name in ['Resumen_Semanal', 'Detalle_OS']:
                                ws = writer.sheets[name]
                                ws.sheet_view.showGridLines = True
                                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                                    for cell in row:
                                        cell.border = borde
                                        if cell.row == 1: cell.fill = fondo; cell.font = blanca; cell.alignment = Alignment(horizontal='center', vertical='center')
                                        else:
                                            if "HECTÁREAS" in str(ws.cell(1, cell.column).value): cell.number_format = '#,##0.00'
                                            elif "($" in str(ws.cell(1, cell.column).value) or "%" in str(ws.cell(1, cell.column).value): cell.number_format = '"$" #,##0'
                                    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col) + 4, 32)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.download_button(label="📥 DESCARGAR INFORME DUAL (EXCEL OFICIAL)", data=buffer_neg.getvalue(), file_name=f"Auditoria_Tarifas_{sim_pista}_{sim_fecha_inicio}_a_{sim_fecha_fin}.xlsx", type="primary", use_container_width=True)

    except Exception as e:
        st.error(f"🚨 Falla crítica en los motores: {e}")
