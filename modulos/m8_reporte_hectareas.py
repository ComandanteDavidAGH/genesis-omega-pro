import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import io

# =================================================================
# 🔌 CONEXIÓN DIRECTA A BÓVEDA DE DATOS (NATIVA E INFALIBLE)
# =================================================================
def obtener_cliente_gspread_m8():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

# =================================================================
# 🚁 MOTOR PRINCIPAL DEL RADAR DE HECTÁREAS Y RENDIMIENTO
# =================================================================
def ejecutar(descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
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
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d'):
            try: return datetime.strptime(str(val).strip(), fmt)
            except: pass
        return None

    def fmt_latino(val):
        try:
            return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(val)

    try:
        with st.spinner("🛰️ Escaneando la Bóveda Maestra y Anclando Flotas a sus Bases..."):
            gc = obtener_cliente_gspread_m8()
            if not gc:
                st.error("🚨 ALERTA ROJA: El conector gspread no pudo inicializarse.")
                return
                
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            datos_brutos = boveda.worksheet("TABLA 1").get_all_values()
            
        if not datos_brutos:
            st.error("🚨 ALERTA ROJA: Conexión establecida pero la pestaña 'TABLA 1' está totalmente vacía.")
            return
        elif len(datos_brutos) <= 5:
            st.warning(f"⚠️ La 'TABLA 1' existe pero solo tiene {len(datos_brutos)} filas.")
            return
            
        if len(datos_brutos) > 5:
            columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "HA_NETAS", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "H_PROPORCIONAL", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_TOTAL_AVION", "TARIFA_HA", "RECARGO_HA", "SUBTOTAL", "COSTO_HORA", "PISTA"]
            
            filas_limpias = [r + [""]*(24 - len(r)) for r in datos_brutos[5:]]
            df_rep = pd.DataFrame([r[:24] for r in filas_limpias], columns=columnas)
            
            df_rep['HA_NETAS'] = df_rep['HA_NETAS'].apply(extraer_numero)
            df_rep['H_PROPORCIONAL'] = df_rep['H_PROPORCIONAL'].apply(extraer_numero)
            df_rep['SEMANA'] = df_rep['SEMANA'].astype(str).str.strip()
            df_rep['PISTA'] = df_rep['PISTA'].astype(str).str.strip().str.upper()
            df_rep['HK'] = df_rep['HK'].astype(str).str.strip().str.upper()
            
            # Anclaje de Base Maestra (Para evitar errores de pista por digitación)
            mask_hk = df_rep['HK'] != ""
            if not df_rep[mask_hk].empty:
                mapa_flota = df_rep[mask_hk].groupby('HK')['PISTA'].agg(lambda x: x.value_counts().index[0] if not x.empty else "").to_dict()
                df_rep.loc[mask_hk, 'PISTA'] = df_rep.loc[mask_hk, 'HK'].map(mapa_flota).fillna(df_rep.loc[mask_hk, 'PISTA'])
            
            df_rep = df_rep[(df_rep['PISTA'] != "") & (df_rep['HA_NETAS'] > 0)]
            df_rep['FECHA_DT'] = df_rep['FECHA'].apply(procesar_fecha_pesada)
            df_rep = df_rep.dropna(subset=['FECHA_DT'])
            
            pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
            min_fecha_real = df_rep['FECHA_DT'].min().date() if not df_rep.empty else datetime.now().date()
            max_fecha_real = df_rep['FECHA_DT'].max().date() if not df_rep.empty else datetime.now().date()
            
            # --- 🎛️ PANEL DE CONTROL ---
            st.markdown("### 🎛️ Centro de Comando y Filtros")
            c1, c2, c3, c4 = st.columns([1.5, 1, 1, 1])
            
            vista_seleccionada = c1.radio(
                "👁️ Seleccione la Vista del Radar:", 
                ["📊 Resumen Gerencial", "📅 Mapa Semanal (Detalle)"], 
                horizontal=True
            )
            
            fecha_sel_ini = c2.date_input("📅 Fecha Inicial:", value=min_fecha_real, key="m8_f_ini_def")
            fecha_sel_fin = c3.date_input("📅 Fecha Final:", value=max_fecha_real, key="m8_f_fin_def")
            pista_sel = c4.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp, key="m8_pista_perfecta")
            
            mostrar_horas = False
            calcular_rend_prom = False
            agrupar_avion = False # 💥 EL INTERRUPTOR 💥

            if vista_seleccionada == "📊 Resumen Gerencial":
                st.markdown("##### ⚙️ Ajustes de Visualización")
                cc1, cc2, cc3 = st.columns(3)
                mostrar_horas = cc1.checkbox("⏱️ Mostrar Horas de Vuelo", value=True)
                calcular_rend_prom = cc2.checkbox("🚀 Mostrar Rend. Promedio (Ha/Hr)", value=True)
                agrupar_avion = cc3.toggle("✈️ Desglosar por Avión (HK)", value=False) # El interruptor solicitado

            df_filt = df_rep[(df_rep['FECHA_DT'].dt.date >= fecha_sel_ini) & (df_rep['FECHA_DT'].dt.date <= fecha_sel_fin)].copy()
            if pista_sel != "TODAS":
                df_filt = df_filt[df_filt['PISTA'] == pista_sel]
            
            if df_filt.empty:
                st.warning("⚠️ No hay operaciones registradas para estos parámetros.")
            else:
                meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
                df_filt['MES'] = df_filt['FECHA_DT'].dt.month.map(meses_nom)
                
                st.markdown("---")
                rango_txt = f"{fecha_sel_ini.strftime('%d/%m/%Y')} al {fecha_sel_fin.strftime('%d/%m/%Y')}"
                
                if vista_seleccionada == "📊 Resumen Gerencial":
                    st.markdown(f"#### 📑 Consolidado {'por Avión' if agrupar_avion else 'General'} ({rango_txt})")
                    
                    tabla_final = []
                    total_hr_gral = 0
                    total_ha_gral = 0

                    if agrupar_avion:
                        # 💥 VISTA DESGLOSADA POR AVIÓN (CON INTERRUPTOR ENCENDIDO) 💥
                        df_gerencia = df_filt.groupby(['PISTA', 'HK', 'MES']).agg(
                            REND_HR=('H_PROPORCIONAL', 'sum'),
                            AREA_FUMIG=('HA_NETAS', 'sum')
                        ).reset_index()
                        
                        for pista in sorted(df_gerencia['PISTA'].unique()):
                            df_pista = df_gerencia[df_gerencia['PISTA'] == pista]
                            sum_hr_pista = df_pista['REND_HR'].sum()
                            sum_ha_pista = df_pista['AREA_FUMIG'].sum()
                            
                            fila_pista = {'NIVEL': f"📍 BASE: {pista}", 'AVIÓN (HK)': '', 'MES': 'TOTAL BASE'}
                            if mostrar_horas or calcular_rend_prom: fila_pista['REND (hr)'] = sum_hr_pista
                            fila_pista['ÁREA FUMIG (ha)'] = sum_ha_pista
                            if calcular_rend_prom: fila_pista['REND. PROMEDIO (Ha/Hr)'] = sum_ha_pista / sum_hr_pista if sum_hr_pista > 0 else 0.0
                            tabla_final.append(fila_pista)
                            
                            for hk in sorted(df_pista['HK'].unique()):
                                datos_hk = df_pista[df_pista['HK'] == hk].sort_values(by='MES')
                                sum_hr_hk = datos_hk['REND_HR'].sum()
                                sum_ha_hk = datos_hk['AREA_FUMIG'].sum()
                                
                                fila_hk = {'NIVEL': '', 'AVIÓN (HK)': f"✈️ AVION: {hk}", 'MES': 'Total Avión'}
                                if mostrar_horas or calcular_rend_prom: fila_hk['REND (hr)'] = sum_hr_hk
                                fila_hk['ÁREA FUMIG (ha)'] = sum_ha_hk
                                if calcular_rend_prom: fila_hk['REND. PROMEDIO (Ha/Hr)'] = sum_ha_hk / sum_hr_hk if sum_hr_hk > 0 else 0.0
                                tabla_final.append(fila_hk)
                                
                                for _, row in datos_hk.iterrows():
                                    mes_limpio = row['MES'].split('-')[1] if '-' in row['MES'] else row['MES']
                                    fila_mes = {'NIVEL': '', 'AVIÓN (HK)': '', 'MES': mes_limpio}
                                    if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = row['REND_HR']
                                    fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                                    if calcular_rend_prom: fila_mes['REND. PROMEDIO (Ha/Hr)'] = row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0.0
                                    tabla_final.append(fila_mes)
                                    
                            total_hr_gral += sum_hr_pista
                            total_ha_gral += sum_ha_pista
                            
                        fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'AVIÓN (HK)': '', 'MES': ''}
                        if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = total_hr_gral
                        fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                        if calcular_rend_prom: fila_tot['REND. PROMEDIO (Ha/Hr)'] = total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0.0
                        tabla_final.append(fila_tot)
                        
                    else:
                        # 💥 VISTA CLÁSICA (CON INTERRUPTOR APAGADO) 💥
                        df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(
                            REND_HR=('H_PROPORCIONAL', 'sum'),
                            AREA_FUMIG=('HA_NETAS', 'sum')
                        ).reset_index()

                        for pista in sorted(df_gerencia['PISTA'].unique()):
                            datos_pista = df_gerencia[df_gerencia['PISTA'] == pista].sort_values(by='MES')
                            sum_hr = datos_pista['REND_HR'].sum()
                            sum_ha = datos_pista['AREA_FUMIG'].sum()
                            
                            fila_sub = {'NIVEL': f"📍 BASE: {pista}", 'MES': 'TOTAL BASE'}
                            if mostrar_horas or calcular_rend_prom: fila_sub['REND (hr)'] = sum_hr
                            fila_sub['ÁREA FUMIG (ha)'] = sum_ha
                            if calcular_rend_prom: fila_sub['REND. PROMEDIO (Ha/Hr)'] = sum_ha / sum_hr if sum_hr > 0 else 0.0
                            tabla_final.append(fila_sub)
                            
                            for _, row in datos_pista.iterrows():
                                mes_limpio = row['MES'].split('-')[1] if '-' in row['MES'] else row['MES']
                                fila_mes = {'NIVEL': '', 'MES': mes_limpio}
                                if mostrar_horas or calcular_rend_prom: fila_mes['REND (hr)'] = row['REND_HR']
                                fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                                if calcular_rend_prom: fila_mes['REND. PROMEDIO (Ha/Hr)'] = row['AREA_FUMIG'] / row['REND_HR'] if row['REND_HR'] > 0 else 0.0
                                tabla_final.append(fila_mes)
                                
                            total_hr_gral += sum_hr
                            total_ha_gral += sum_ha
                            
                        fila_tot = {'NIVEL': '👑 TOTAL GENERAL', 'MES': ''}
                        if mostrar_horas or calcular_rend_prom: fila_tot['REND (hr)'] = total_hr_gral
                        fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                        if calcular_rend_prom: fila_tot['REND. PROMEDIO (Ha/Hr)'] = total_ha_gral / total_hr_gral if total_hr_gral > 0 else 0.0
                        tabla_final.append(fila_tot)

                    df_visual = pd.DataFrame(tabla_final)
                    
                    def estilizar_filas(row):
                        if "BASE:" in str(row['NIVEL']): return ['background-color: #d1ecf1; font-weight: bold; color: #0c5460;'] * len(row)
                        elif "TOTAL GENERAL" in str(row['NIVEL']): return ['background-color: #c3e6cb; font-weight: bold; color: #155724;'] * len(row)
                        elif 'AVIÓN (HK)' in row and "✈️" in str(row['AVIÓN (HK)']): return ['background-color: #f8f9fa; font-weight: bold; color: #212529;'] * len(row)
                        return [''] * len(row)
                    
                    formato_columnas = {'ÁREA FUMIG (ha)': fmt_latino}
                    if mostrar_horas or calcular_rend_prom: formato_columnas['REND (hr)'] = fmt_latino
                    if calcular_rend_prom: formato_columnas['REND. PROMEDIO (Ha/Hr)'] = fmt_latino
                    
                    st.dataframe(
                        df_visual.style.apply(estilizar_filas, axis=1).format(formato_columnas),
                        use_container_width=True,
                        hide_index=True
                    )

                else:
                    matriz = pd.pivot_table(df_filt, values='HA_NETAS', index='MES', columns='SEMANA', aggfunc='sum', fill_value=0)
                    matriz = matriz.sort_index()
                    cols_ordenadas = sorted(matriz.columns, key=lambda x: int(x) if str(x).isdigit() else 999)
                    matriz = matriz[cols_ordenadas]
                    
                    matriz.index = [m.split('-')[1] if '-' in m else m for m in matriz.index]
                    matriz['TOTAL MES'] = matriz.sum(axis=1)
                    matriz.loc['TOTAL ANUAL'] = matriz.sum(axis=0)
                    
                    st.markdown(f"#### 🛩️ Rendimiento Semana a Semana: **{pista_sel}** ({rango_txt})")
                    
                    if HAS_MATPLOTLIB:
                        st.dataframe(matriz.style.format(fmt_latino).background_gradient(cmap="YlGn", axis=None), use_container_width=True)
                    else:
                        st.dataframe(matriz.style.format(fmt_latino), use_container_width=True)
                    
                    st.markdown("---")
                    df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
                    if not df_grafico.empty:
                        df_grafico['TEXTO_LATINO'] = df_grafico['TOTAL MES'].apply(fmt_latino)
                        
                        fig = px.bar(
                            df_grafico, x='index', y='TOTAL MES', text='TEXTO_LATINO',
                            labels={'TOTAL MES': 'Hectáreas Fumigadas', 'index': 'Mes de Operación'},
                            color='TOTAL MES', color_continuous_scale='Greens'
                        )
                        fig.update_traces(textposition='outside')
                        fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', showlegend=False, xaxis_title="Mes", separators=",.")
                        st.plotly_chart(fig, use_container_width=True)

                st.markdown("---")
                buffer_rep = io.BytesIO()
                with pd.ExcelWriter(buffer_rep, engine='openpyxl') as writer:
                    nombre_hoja = 'Resumen_Gerencial' if "Gerencial" in vista_seleccionada else 'Reporte_Semanal'
                    
                    if "Gerencial" in vista_seleccionada:
                        df_visual.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    else:
                        matriz.to_excel(writer, sheet_name=nombre_hoja)
                        
                    workbook = writer.book
                    worksheet = writer.sheets[nombre_hoja]
                    worksheet.sheet_view.showGridLines = False
                    worksheet.row_dimensions[1].height = 30
                    
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.chart import BarChart, Reference
                    from openpyxl.utils import get_column_letter

                    borde_pro = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), 
                                       top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                    navy_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                    white_font = Font(color="FFFFFF", bold=True, size=11)
                    months_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
                    pista_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid") 
                    sub_fill = PatternFill(start_color="E2E6EA", end_color="E2E6EA", fill_type="solid") 
                    total_fill = PatternFill(start_color="C3E6CB", end_color="C3E6CB", fill_type="solid") 

                    max_row = worksheet.max_row
                    max_col = worksheet.max_column

                    for row in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                        for cell in row:
                            cell.border = borde_pro
                            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.number_format = '#,##0.00'
                            if cell.row == 1:
                                cell.fill = navy_fill; cell.font = white_font
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            else:
                                cell.alignment = Alignment(vertical='center', indent=1)
                                
                            if "Gerencial" in vista_seleccionada and cell.row > 1:
                                nivel_v = str(worksheet.cell(row=cell.row, column=1).value or "").strip()
                                if "BASE:" in nivel_v: cell.fill = pista_fill; cell.font = Font(bold=True, color="0C5460")
                                elif "TOTAL GENERAL" in nivel_v: cell.fill = total_fill; cell.font = Font(bold=True, color="155724")
                                elif nivel_v == "" or nivel_v == "None": cell.fill = months_fill
                                
                                if agrupar_avion:
                                    avion_v = str(worksheet.cell(row=cell.row, column=2).value or "").strip()
                                    if "✈️" in avion_v: cell.fill = sub_fill; cell.font = Font(bold=True)

                    chart = BarChart()
                    chart.type = "col"; chart.style = 10
                    chart.title = "Rendimiento Operativo (Ha)"; chart.y_axis.title = "Hectáreas"
                    
                    if "Gerencial" in vista_seleccionada:
                        idx_ha = df_visual.columns.get_loc('ÁREA FUMIG (ha)') + 1
                        col_ha_letra = get_column_letter(idx_ha)
                        
                        col_lbl_chart = max_col + 2
                        col_val_chart = max_col + 3
                        
                        worksheet.cell(row=1, column=col_lbl_chart).value = "Etiqueta_Grafico"
                        worksheet.cell(row=1, column=col_val_chart).value = "Ha"
                        row_g = 2
                        
                        for r_b in range(2, max_row + 1):
                            n_v = str(worksheet.cell(row=r_b, column=1).value or "")
                            
                            if agrupar_avion:
                                a_v = str(worksheet.cell(row=r_b, column=2).value or "")
                                m_v = str(worksheet.cell(row=r_b, column=3).value or "")
                                if n_v == "" and a_v == "" and m_v not in ["Total Avión", "TOTAL BASE", ""]:
                                    av_encontrado = "Desc"
                                    for r_back in range(r_b, 1, -1):
                                        if "✈️" in str(worksheet.cell(row=r_back, column=2).value):
                                            av_encontrado = str(worksheet.cell(row=r_back, column=2).value).replace("✈️ AVION:", "").strip()
                                            break
                                    worksheet.cell(row=row_g, column=col_lbl_chart).value = f"{m_v} ({av_encontrado})"
                                    worksheet.cell(row=row_g, column=col_val_chart).value = f"={col_ha_letra}{r_b}"
                                    row_g += 1
                            else:
                                m_v = str(worksheet.cell(row=r_b, column=2).value or "")
                                if n_v == "" and m_v not in ["TOTAL BASE", ""]:
                                    pista_encontrada = "Desc"
                                    for r_back in range(r_b, 1, -1):
                                        if "BASE:" in str(worksheet.cell(row=r_back, column=1).value):
                                            pista_encontrada = str(worksheet.cell(row=r_back, column=1).value).replace("📍 BASE:", "").strip()
                                            break
                                    worksheet.cell(row=row_g, column=col_lbl_chart).value = f"{m_v} ({pista_encontrada})"
                                    worksheet.cell(row=row_g, column=col_val_chart).value = f"={col_ha_letra}{r_b}"
                                    row_g += 1
                        
                        if row_g > 2:
                            data = Reference(worksheet, min_col=col_val_chart, min_row=1, max_row=row_g-1)
                            cats = Reference(worksheet, min_col=col_lbl_chart, min_row=2, max_row=row_g-1)
                            chart.add_data(data, titles_from_data=True)
                            chart.set_categories(cats)
                            
                            for r_inv in range(1, row_g):
                                worksheet.cell(row=r_inv, column=col_lbl_chart).font = Font(color="FFFFFF")
                                worksheet.cell(row=r_inv, column=col_val_chart).font = Font(color="FFFFFF")
                                
                            worksheet.add_chart(chart, f"{get_column_letter(max_col + 1)}2")
                    else:
                        data = Reference(worksheet, min_col=max_col, min_row=1, max_row=max_row-1)
                        cats = Reference(worksheet, min_col=1, min_row=2, max_row=max_row-1)
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        worksheet.add_chart(chart, f"{get_column_letter(max_col + 2)}2")
                    
                    for col_idx in range(1, max_col + 1):
                        worksheet.column_dimensions[get_column_letter(col_idx)].width = 24
                    worksheet.freeze_panes = "A2"

                rango_label = f"{fecha_sel_ini.strftime('%Y%m%d')}_{fecha_sel_fin.strftime('%Y%m%d')}"
                st.download_button(
                    label="💾 DESCARGAR REPORTE GERENCIAL TOP",
                    data=buffer_rep.getvalue(),
                    file_name=f"Reporte_Rendimiento_{rango_label}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )                       
    except Exception as e:
        st.error(f"🚨 Falla en el sistema de radares: {e}")
