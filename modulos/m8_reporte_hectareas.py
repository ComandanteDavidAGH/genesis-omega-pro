import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

# =================================================================
# ⚡ MOTORES DE CACHÉ Y COMPUTACIÓN ACELERADA (ALTA VELOCIDAD)
# =================================================================

@st.cache_data(show_spinner=False)
def cargar_y_preprocesar_base_radar(_descargar_matriz_rapida, _procesar_fecha_pesada, _extraer_numero):
    """ Descarga, estructura y decodifica la base histórica una sola vez en RAM """
    url_maestra = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    datos_brutos = _descargar_matriz_rapida(url_maestra, "TABLA 1")
    
    if len(datos_brutos) <= 5:
        return pd.DataFrame()
        
    columnas = ["OS", "BLOQUE", "FINCA", "SECTOR", "AREA_BRUTA", "HA_NETAS", "COCTEL", "FECHA", "DIA", "SEMANA", "H_TOTAL", "GLN_HA", "VOL_TOTAL", "H_PROPORCIONAL", "REND_MIN", "PILOTO", "HK", "MODELO", "COSTO_TOTAL_AVION", "TARIFA_HA", "RECARGO_HA", "SUBTOTAL", "COSTO_HORA", "PISTA"]
    
    filas_limpias = [r + [""]*(24 - len(r)) for r in datos_brutos[5:]]
    df = pd.DataFrame([r[:24] for r in filas_limpias], columns=columnas)
    
    df['HA_NETAS'] = df['HA_NETAS'].apply(_extraer_numero)
    df['H_PROPORCIONAL'] = df['H_PROPORCIONAL'].apply(_extraer_numero)
    df['SEMANA'] = df['SEMANA'].astype(str).str.strip()
    df['PISTA'] = df['PISTA'].astype(str).str.strip().str.upper()
    
    df = df[(df['PISTA'] != "") & (df['HA_NETAS'] > 0)]
    meses_nom = {1:"01-ene", 2:"02-feb", 3:"03-mar", 4:"04-abr", 5:"05-may", 6:"06-jun", 7:"07-jul", 8:"08-ago", 9:"09-sep", 10:"10-oct", 11:"11-nov", 12:"12-dic"}
    
    # Mapeo temporal acelerado
    lista_meses, lista_anios = [], []
    for f_str in df['FECHA']:
        dt = _procesar_fecha_pesada(f_str)
        if dt:
            lista_meses.append(meses_nom.get(dt.month, "00-Desc"))
            lista_anios.append(str(dt.year))
        else:
            lista_meses.append("00-Desc")
            lista_anios.append("00-Desc")
            
    df['MES'] = lista_meses
    df['AÑO'] = lista_anios
    return df[df['AÑO'] != "00-Desc"].reset_index(drop=True)


def compilar_excel_radar_on_demand(df_visual, matriz, vista, mostrar_horas, anio_sel, pista_sel, col_ha_letra, col_ha_idx):
    """ 🚀 LAZY COMPILATION: openpyxl solo consume CPU en el momento de la descarga """
    buffer_rep = io.BytesIO()
    nombre_hoja = 'Resumen_Gerencial' if "Gerencial" in vista else 'Reporte_Semanal'
    
    with pd.ExcelWriter(buffer_rep, engine='openpyxl') as writer:
        if "Gerencial" in vista:
            df_visual.to_excel(writer, sheet_name=nombre_hoja, index=False)
        else:
            matriz.to_excel(writer, sheet_name=nombre_hoja)
            
        worksheet = writer.sheets[nombre_hoja]
        worksheet.sheet_view.showGridLines = True  # Corrección: Forzar rejilla visible
        worksheet.row_dimensions[1].height = 30
        
        borde_pro = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), 
                           top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
        fondo_navy = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        fuente_blanca = Font(color="FFFFFF", bold=True, size=11)
        fondo_meses = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
        fondo_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        fondo_total = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")

        max_row = worksheet.max_row
        max_col = worksheet.max_column

        if "Gerencial" in vista:
            rango_total_ha = []
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
                    
                if "Gerencial" in vista and cell.row > 1:
                    nivel_v = str(worksheet.cell(row=cell.row, column=1).value or "").strip()
                    if "➖" in nivel_v:
                        cell.fill = fondo_sub; cell.font = Font(bold=True)
                    elif "TOTAL GENERAL" in nivel_v:
                        cell.fill = fondo_total; cell.font = Font(bold=True, color="FFFFFF")
                    elif nivel_v == "" or nivel_v == "None":
                        cell.fill = fondo_meses

        chart = BarChart()
        chart.type = "col"; chart.style = 10
        chart.title = f"Rendimiento Operativo (Ha) - Base {pista_sel}"; chart.y_axis.title = "Hectáreas"
        chart.legend = None
        chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True
        chart.height = 14; chart.width = 24
        
        if "Gerencial" in vista:
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
        
    return buffer_rep.getvalue()

# =================================================================
# 👑 INTERFAZ GRÁFICA Y SEGMENTACIÓN DE VISTAS
# =================================================================

def ejecutar(descargar_matriz_rapida, extraer_numero, procesar_fecha_pesada, HAS_MATPLOTLIB):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"], div[data-testid="stDataFrame"] > div { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    
    /* HUD de Control de Rendimiento */
    .hud-radar {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .hud-radar-item { text-align: center; flex: 1; }
    .hud-radar-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-radar-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Radar de Hectáreas y Rendimiento</h1>", unsafe_allow_html=True)
    
    # ⚡ EXTRACCIÓN MAESTRA CACHEADA EN RAM
    df_rep = cargar_y_preprocesar_base_radar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero)
    
    if df_rep.empty:
        st.warning("⚠️ Bóveda vacía o sin misiones activas registradas en la TABLA 1.")
        return

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

    # Filtrado lógico veloz en memoria RAM local
    df_filt = df_rep[df_rep['AÑO'] == año_sel]
    if pista_sel != "TODAS":
        df_filt = df_filt[df_filt['PISTA'] == pista_sel]
    
    if df_filt.empty:
        st.warning("⚠️ No hay operaciones registradas para los parámetros elegidos.")
    else:
        # 🚀 LANZAMIENTO DEL HUD GERENCIAL DE ALTO IMPACTO
        total_ha_filtro = df_filt['HA_NETAS'].sum()
        total_hr_filtro = df_filt['H_PROPORCIONAL'].sum()
        ratio_eficiencia = total_ha_filtro / total_hr_filtro if total_hr_filtro > 0 else 0.0

        hb1, hb2, hb3 = st.columns(3)
        with hb1: st.markdown(f"<div class='hud-radar'><div class='hud-radar-item'><p class='hud-radar-title'>Volumen de Operación</p><p class='hud-radar-value'>🚜 {total_ha_filtro:,.2f} Ha</p></div></div>", unsafe_allow_html=True)
        with hb2: st.markdown(f"<div class='hud-radar'><div class='hud-radar-item'><p class='hud-radar-title'>Horas Totales del Aire</p><p class='hud-radar-value'>⏱️ {total_hr_filtro:,.2f} Hr</p></div></div>", unsafe_allow_html=True)
        with hb3: st.markdown(f"<div class='hud-radar'><div class='hud-radar-item'><p class='hud-radar-title'>Ratio de Rendimiento</p><p class='hud-radar-value'>🛰️ {ratio_eficiencia:,.1f} Ha/Hr</p></div></div>", unsafe_allow_html=True)

        st.markdown("---")
        df_visual = pd.DataFrame()
        matriz = pd.DataFrame()
        
        col_ha_letra = "C" if not mostrar_horas else "D"
        col_ha_idx = 3 if not mostrar_horas else 4
        
        if vista_seleccionada == "📊 Resumen Gerencial (Hectáreas)":
            st.markdown(f"#### 📑 Consolidado Gerencial - {año_sel}")
            
            df_gerencia = df_filt.groupby(['PISTA', 'MES']).agg(
                REND_HR=('H_PROPORCIONAL', 'sum'),
                AREA_FUMIG=('HA_NETAS', 'sum')
            ).reset_index()
            
            tabla_final = []
            total_hr_gral, total_ha_gral = 0.0, 0.0
            
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
                if "➖" in str(row['NIVEL']) or "TOTAL" in str(row['NIVEL']):
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
        
        # ⚡ MANIOBRA LAZY DE ALTA VELOCIDAD: openpyxl solo compila si se solicita la descarga
        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL TOP (EXCEL)",
            data=compilar_excel_radar_on_demand(df_visual, matriz, vista_seleccionada, mostrar_horas, año_sel, pista_sel, col_ha_letra, col_ha_idx),
            file_name=f"Reporte_Gerencial_Rendimiento_{año_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
