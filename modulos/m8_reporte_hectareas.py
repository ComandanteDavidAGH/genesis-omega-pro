import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

def ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada, HAS_MATPLOTLIB):
    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
    try:
        with st.spinner("🛰️ Escaneando la Bóveda Maestra con Motor Turbo (TABLA 1)..."):
            url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
            datos_brutos = descargar_matriz_rapida(url_maestra, "TABLA 1")
            
        # =================================================================
        # 🚨 TRAMPA DE DIAGNÓSTICO (Para evitar pantallas blancas)
        # =================================================================
        if not datos_brutos:
            st.error("🚨 ALERTA ROJA: Falla de conexión con Google Sheets. La Bóveda (TABLA 1) no envió información.")
            return
        elif len(datos_brutos) <= 5:
            st.warning(f"⚠️ La TABLA 1 tiene solo {len(datos_brutos)} filas. Se requieren más de 5 filas para encender el radar.")
            return
        # =================================================================
            
        if len(datos_brutos) > 5:
            columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "HA_NETAS", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "H_PROPORCIONAL", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_TOTAL_AVION", "TARIFA_HA", "RECARGO_HA", "SUBTOTAL", "COSTO_HORA", "PISTA"]
            
            filas_limpias = [r + [""]*(24 - len(r)) for r in datos_brutos[5:]]
            df_rep = pd.DataFrame([r[:24] for r in filas_limpias], columns=columnas)
            
            df_rep['HA_NETAS'] = df_rep['HA_NETAS'].apply(extraer_numero)
            df_rep['H_PROPORCIONAL'] = df_rep['H_PROPORCIONAL'].apply(extraer_numero)
            df_rep['SEMANA'] = df_rep['SEMANA'].astype(str).str.strip()
            df_rep['PISTA'] = df_rep['PISTA'].astype(str).str.strip().str.upper()
            
            df_rep = df_rep[(df_rep['PISTA'] != "") & (df_rep['HA_NETAS'] > 0)]
            
            meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
            
            def extraer_mes_año(fecha_str):
                dt = procesar_fecha_pesada(fecha_str)
                if dt: return meses_nom.get(dt.month, "00-Desc"), str(dt.year)
                return "00-Desc", "00-Desc"
            
            df_rep[['MES', 'AÑO']] = df_rep['FECHA'].apply(lambda x: pd.Series(extraer_mes_año(x)))
            df_rep = df_rep[df_rep['AÑO'] != "00-Desc"]
            
            st.markdown("### 🎛️ Centro de Comando y Filtros")
            c1, c2, c3 = st.columns([2, 1, 1])
            
            vista_seleccionada = c1.radio(
                "👁️ Seleccione la Vista del Radar:", 
                ["📊 Resumen Gerencial (Hectáreas)", "📅 Mapa Semanal (Detalle)"], 
                horizontal=True
            )
            
            pistas_disp = sorted(df_rep['PISTA'].unique().tolist())
            años_disp = sorted(df_rep['AÑO'].unique().tolist(), reverse=True)
            
            año_sel = c2.selectbox("📅 Año Fiscal", años_disp if años_disp else [str(datetime.now().year)])
            pista_sel = c3.selectbox("📍 Base (Pista)", ["TODAS"] + pistas_disp)
            
            mostrar_horas = False
            if vista_seleccionada == "📊 Resumen Gerencial (Hectáreas)":
                mostrar_horas = st.checkbox("⏱️ Mostrar también el Rendimiento (Horas de Vuelo)")

            df_filt = df_rep[df_rep['AÑO'] == año_sel]
            if pista_sel != "TODAS":
                df_filt = df_filt[df_filt['PISTA'] == pista_sel]
            
            if df_filt.empty:
                st.warning("⚠️ No hay operaciones registradas para estos parámetros.")
            else:
                st.markdown("---")
                
                if vista_seleccionada == "📊 Resumen Gerencial (Hectáreas)":
                    st.markdown(f"#### 📑 Consolidado Gerencial - {año_sel}")
                    
                    df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(
                        REND_HR=('H_PROPORCIONAL', 'sum'),
                        AREA_FUMIG=('HA_NETAS', 'sum')
                    ).reset_index()
                    
                    tabla_final = []
                    total_hr_gral = 0
                    total_ha_gral = 0
                    
                    for pista in sorted(df_gerencia['PISTA'].unique()):
                        datos_pista = df_gerencia[df_gerencia['PISTA'] == pista].sort_values(by='MES')
                        sum_hr = datos_pista['REND_HR'].sum()
                        sum_ha = datos_pista['AREA_FUMIG'].sum()
                        
                        fila_sub = {'NIVEL': f"➖ {pista}", 'MES': ''}
                        if mostrar_horas: fila_sub['REND (hr)'] = sum_hr
                        fila_sub['ÁREA FUMIG (ha)'] = sum_ha
                        tabla_final.append(fila_sub)
                        
                        for _, row in datos_pista.iterrows():
                            mes_limpio = row['MES'].split('-')[1] if '-' in row['MES'] else row['MES']
                            fila_mes = {'NIVEL': '', 'MES': mes_limpio}
                            if mostrar_horas: fila_mes['REND (hr)'] = row['REND_HR']
                            fila_mes['ÁREA FUMIG (ha)'] = row['AREA_FUMIG']
                            tabla_final.append(fila_mes)
                            
                        total_hr_gral += sum_hr
                        total_ha_gral += sum_ha
                        
                    fila_tot = {'NIVEL': 'TOTAL GENERAL', 'MES': ''}
                    if mostrar_horas: fila_tot['REND (hr)'] = total_hr_gral
                    fila_tot['ÁREA FUMIG (ha)'] = total_ha_gral
                    tabla_final.append(fila_tot)
                    
                    df_visual = pd.DataFrame(tabla_final)
                    
                    def estilizar_filas(row):
                        if "➖" in row['NIVEL'] or "TOTAL" in row['NIVEL']:
                            return ['background-color: #e2e6ea; font-weight: bold;'] * len(row)
                        return [''] * len(row)
                    
                    formato_columnas = {'ÁREA FUMIG (ha)': "{:.2f}"}
                    if mostrar_horas: formato_columnas['REND (hr)'] = "{:.2f}"
                    
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
                    
                    st.markdown(f"#### 🚜 Rendimiento Semana a Semana: **{pista_sel}**")
                    if HAS_MATPLOTLIB:
                        st.dataframe(matriz.style.format("{:.2f}").background_gradient(cmap="YlGn", axis=None), use_container_width=True)
                    else:
                        st.dataframe(matriz.style.format("{:.2f}"), use_container_width=True)
                    
                    st.markdown("---")
                    df_grafico = matriz.drop('TOTAL ANUAL', errors='ignore').reset_index()
                    if not df_grafico.empty:
                        fig = px.bar(
                            df_grafico, x='index', y='TOTAL MES', text='TOTAL MES',
                            labels={'TOTAL MES': 'Hectáreas Fumigadas', 'index': 'Mes de Operación'},
                            color='TOTAL MES', color_continuous_scale='Greens'
                        )
                        fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                        fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', showlegend=False, xaxis_title="Mes")
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
                    from openpyxl.chart.label import DataLabelList
                    from openpyxl.utils import get_column_letter

                    borde_pro = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), 
                                       top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                    fondo_navy = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                    fuente_blanca = Font(color="FFFFFF", bold=True, size=11)
                    fondo_meses = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
                    fondo_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    fondo_total = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")

                    max_row = worksheet.max_row
                    max_col = worksheet.max_column

                    if "Gerencial" in vista_seleccionada:
                        rango_total_ha = []
                        col_ha_letra = "C" if not mostrar_horas else "D"
                        col_ha_idx = 3 if not mostrar_horas else 4
                        
                        for i in range(2, max_row + 1):
                            nivel = str(worksheet.cell(row=i, column=1).value or "").strip()
                            if "➖" in nivel:
                                inicio = i + 1
                                fin = i + 1
                                for j in range(i + 1, max_row + 1):
                                    val_j = str(worksheet.cell(row=j, column=1).value or "").strip()
                                    if val_j == "" or val_j == "None": fin = j
                                    else: break
                                worksheet.cell(row=i, column=col_ha_idx).value = f"=SUM({col_ha_letra}{inicio}:{col_ha_letra}{fin})"
                                rango_total_ha.append(f"{col_ha_letra}{i}")
                            elif "TOTAL GENERAL" in nivel:
                                if rango_total_ha:
                                    worksheet.cell(row=i, column=col_ha_idx).value = f"=SUM({','.join(rango_total_ha)})"

                    for row in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                        for cell in row:
                            cell.border = borde_pro
                            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith('=')):
                                cell.number_format = '#,##0.00'
                            if cell.row == 1:
                                cell.fill = fondo_navy; cell.font = fuente_blanca
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            else:
                                cell.alignment = Alignment(vertical='center', indent=1)
                                
                            if "Gerencial" in vista_seleccionada and cell.row > 1:
                                nivel_v = str(worksheet.cell(row=cell.row, column=1).value or "").strip()
                                if "➖" in nivel_v:
                                    cell.fill = fondo_sub; cell.font = Font(bold=True)
                                elif "TOTAL GENERAL" in nivel_v:
                                    cell.fill = fondo_total; cell.font = Font(bold=True, color="FFFFFF")
                                elif nivel_v == "" or nivel_v == "None":
                                    cell.fill = fondo_meses

                    chart = BarChart()
                    chart.type = "col"; chart.style = 10
                    chart.title = "Rendimiento Operativo (Ha)"; chart.y_axis.title = "Hectáreas"
                    chart.legend = None
                    chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True
                    chart.height = 14; chart.width = 24
                    
                    if "Gerencial" in vista_seleccionada:
                        worksheet.cell(row=1, column=27).value = "Mes"
                        worksheet.cell(row=1, column=28).value = "Ha"
                        meses_para_grafico = [m for m in df_visual['MES'] if str(m).strip() not in ["", "None"]]
                        row_g = 2
                        for m in meses_para_grafico:
                            worksheet.cell(row=row_g, column=27).value = m
                            fila_origen = 2
                            for r_b in range(2, max_row):
                                if str(worksheet.cell(row=r_b, column=2).value) == m:
                                    fila_origen = r_b; break
                            worksheet.cell(row=row_g, column=28).value = f"={col_ha_letra}{fila_origen}"
                            row_g += 1
                        
                        data = Reference(worksheet, min_col=28, min_row=1, max_row=row_g-1)
                        cats = Reference(worksheet, min_col=27, min_row=2, max_row=row_g-1)
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        
                        for r_inv in range(1, row_g):
                            worksheet.cell(row=r_inv, column=27).font = Font(color="FFFFFF")
                            worksheet.cell(row=r_inv, column=28).font = Font(color="FFFFFF")
                        
                        worksheet.add_chart(chart, "H2")
                    else:
                        data = Reference(worksheet, min_col=max_col, min_row=1, max_row=max_row-1)
                        cats = Reference(worksheet, min_col=1, min_row=2, max_row=max_row-1)
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        worksheet.add_chart(chart, f"{get_column_letter(max_col + 2)}2")
                    
                    for col_idx in range(1, max_col + 1):
                        worksheet.column_dimensions[get_column_letter(col_idx)].width = 22
                    worksheet.freeze_panes = "A2"

                st.download_button(
                    label="💾 DESCARGAR REPORTE GERENCIAL TOP",
                    data=buffer_rep.getvalue(),
                    file_name=f"Reporte_Gerencial_{año_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )                       
    except Exception as e:
        st.error(f"🚨 Falla en el sistema de radares: {e}")
