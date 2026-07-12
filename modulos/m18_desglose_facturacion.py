import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import gspread
import re
import math
import io
import openpyxl
from datetime import datetime, date
from oauth2client.service_account import ServiceAccountCredentials

# 🛰️ ENLACES NATIVOS
from modulos.utilidades import procesar_fecha_pesada
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =================================================================
# 🔌 CONEXIÓN Y MOTORES DE FORMATO REGIONAL BLINDADOS
# =================================================================

def formato_latino(numero, decimales=0):
    if pd.isna(numero) or numero is None: return "0"
    try:
        num = float(numero)
        if num == 0: return "0"
        if decimales == 0: texto_us = f"{num:,.0f}"
        else: texto_us = f"{num:,.{decimales}f}"
        return texto_us.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0"

def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

# 💥 TRANSLATOR PRO: Evita el colapso por puntos múltiples repetidos de SAP
def limpiar_tarifa_excel(val):
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip().replace("$", "").replace(" ", "").upper()
    if not v or v in ['-', 'NAN', 'NONE', '']: return 0.0
    
    s_clean = re.sub(r'[^\d\.,\-]', '', v)
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
    except:
        return 0.0

def normalizar_a_fecha_pura(val):
    try:
        res_nativo = procesar_fecha_pesada(val)
        if isinstance(res_nativo, (datetime, pd.Timestamp)): return res_nativo.date()
        if isinstance(res_nativo, date): return res_nativo
        return pd.to_datetime(str(res_nativo)).date()
    except: return None

@st.cache_data(show_spinner=False, ttl=600)
def cargar_bases_m18():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try:
        boveda_act = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        
        # 1. CARGAR TABLA 1 (HISTÓRICO)
        datos_brutos = boveda_act.worksheet("TABLA 1").get_all_values()
        df_t1 = pd.DataFrame()
        if len(datos_brutos) > 5:
            idx_t1 = 4
            for i in range(min(8, len(datos_brutos))):
                if "FINCA" in [str(x).upper().strip() for x in datos_brutos[i]]:
                    idx_t1 = i
                    break
            
            encabezados = [str(c).upper().strip().replace("\n", "") for c in datos_brutos[idx_t1]]
            filas_limpias = [r + [""]*(len(encabezados) - len(r)) for r in datos_brutos[idx_t1+1:]]
            df_t1 = pd.DataFrame([r[:len(encabezados)] for r in filas_limpias], columns=encabezados)

        # 2. CARGAR TABLA 2 (PRODUCTORES)
        df_t2 = pd.DataFrame()
        try: 
            t2_raw = boveda_act.worksheet("TABLA 2").get_all_values()
            idx_t2 = next((i for i, r in enumerate(t2_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 0)
            df_t2 = pd.DataFrame(t2_raw[idx_t2+1:], columns=[str(c).strip().upper() for c in t2_raw[idx_t2]])
        except: pass

        # 3. CARGAR CONFIGURACIÓN (TARIFAS ST)
        df_cfg = pd.DataFrame()
        try:
            cfg_raw = boveda_act.worksheet("Configuración").get_all_values()
            df_cfg = pd.DataFrame(cfg_raw[1:], columns=[str(c).upper().strip() for c in cfg_raw[0]])
        except: pass

        return df_t1, df_t2, df_cfg
    except Exception as e:
        raise Exception(f"Error de conexión con Google Drive: {e}")

# =================================================================
# 👑 RENDERIZADO VISUAL PRINCIPAL
# =================================================================

def ejecutar(*args, **kwargs):
    VERDE_INTENSO = '#143521'

    # 💥 BLOQUE CSS SANEADO (Sin f-strings para evitar SyntaxError)
    css_maestro = """
    <style>
    .titulo-desglose { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; margin-bottom: 15px;}
    div[data-testid="stDataFrame"] { border: 3px solid VERDE_HEX !important; border-radius: 8px !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); overflow: hidden !important; }
    .tarjeta-kpi { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.2); text-align: center; margin-bottom: 15px;}
    .kpi-titulo { font-size: 12px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .kpi-valor { font-size: 24px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    
    div[data-testid="stDateInput"] input,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] {
        background-color: #ffffff !important;
        border: 3px solid VERDE_HEX !important;
        border-radius: 6px !important;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
        background-color: transparent !important;
        border: none !important;
    }
    div[data-testid="stDateInput"] *, div[data-testid="stMultiSelect"] * {
        color: #000000 !important;
        font-weight: bold !important;
    }
    div[data-testid="stMainBlockContainer"] label p {
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }
    </style>
    """.replace("VERDE_HEX", VERDE_INTENSO)
    
    st.markdown(css_maestro, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-desglose'>🔍 Módulo 18: Desglose y Auditoría de Facturación</h1>", unsafe_allow_html=True)
    st.write("Ingeniería inversa sobre la TABLA 1: Desglosa facturas ya ejecutadas para revelar el costo real de la mezcla química.")

    with st.spinner("Conectando con la Bóveda Maestra y calculando intervalos..."):
        # 💥 SOLUCIÓN APLICADA: Ahora llama correctamente a cargar_bases_m18()
        df_t1, df_t2, df_cfg = cargar_bases_m18()

    if df_t1.empty:
        st.error("🚨 La Bóveda de Datos (TABLA 1) está vacía o inaccesible.")
        return

    # 1. IDENTIFICACIÓN DE COLUMNAS CLAVE
    col_finca = next((c for c in df_t1.columns if 'FINCA' in c or 'PROPIEDAD' in c), None)
    col_fecha = next((c for c in df_t1.columns if 'FECHA' in c or 'DATE' in c), None)
    col_ha = next((c for c in df_t1.columns if 'FUMIG' in c or 'NETA' in c or 'HECT' in c), None)
    col_coctel = next((c for c in df_t1.columns if 'COCTEL' in c or 'CÓCTEL' in c), None)
    col_precio_vuelo = next((c for c in df_t1.columns if 'COSTO' in c and 'AVI' in c and 'HA' in c.replace(" ", "")), None)
    col_costo_vuelo_finca = next((c for c in df_t1.columns if 'COSTO' in c and 'AVI' in c and 'FINCA' in c), None)
    col_valor_fact = next((c for c in df_t1.columns if 'VALOR' in c and 'FACTURAR' in c), None)

    if not all([col_finca, col_fecha, col_ha, col_precio_vuelo, col_costo_vuelo_finca, col_valor_fact]):
        st.error("🚨 Faltan columnas críticas en la TABLA 1. Verifique los nombres (FINCA, FECHA, ÁREA FUMIG, COSTO AVIÓN ($/ha), COSTO AVIÓN ($/finca), VALOR A FACTURAR).")
        return

    # 2. PREPROCESAMIENTO Y CÁLCULO DE DÍAS CICLO
    df_t1['FECHA_PURA'] = df_t1[col_fecha].apply(normalizar_a_fecha_pura)
    df_t1 = df_t1.dropna(subset=['FECHA_PURA']).copy()
    df_t1['FINCA_CLEAN'] = df_t1[col_finca].astype(str).apply(lambda x: re.sub(r'[^A-Z0-9]', '', x.upper().strip()))
    
    # Ordenar chronológicamente para calcular ciclos
    df_t1 = df_t1.sort_values(by=['FINCA_CLEAN', 'FECHA_PURA'])
    df_t1['FECHA_PREVIA'] = df_t1.groupby('FINCA_CLEAN')['FECHA_PURA'].shift(1)
    
    # Calcular días ciclo (Si es el primer vuelo, asume 14 días por defecto)
    df_t1['DIAS_CICLO'] = (pd.to_datetime(df_t1['FECHA_PURA']) - pd.to_datetime(df_t1['FECHA_PREVIA'])).dt.days
    df_t1['DIAS_CICLO'] = df_t1['DIAS_CICLO'].fillna(14).astype(int)

    # 3. INTERFAZ DE FILTROS
    with st.container(border=True):
        st.markdown("#### 🎛️ Rango de Búsqueda y Selección")
        c1, c2 = st.columns(2)
        fecha_ini = c1.date_input("📅 Fecha Inicial", value=date(2026, 7, 1))
        fecha_fin = c2.date_input("📆 Fecha Final", value=date.today())

        # Filtrar solo para mostrar las fincas operadas en ese rango
        mask_fechas = (df_t1['FECHA_PURA'] >= fecha_ini) & (df_t1['FECHA_PURA'] <= fecha_fin)
        fincas_disponibles = sorted(df_t1[mask_fechas][col_finca].astype(str).str.upper().str.strip().unique().tolist())
        
        fincas_sel = st.multiselect("📍 Seleccione las Fincas a desglosar (Deje vacío para analizarlas TODAS):", fincas_disponibles)

    if st.button("🔥 EJECUTAR DESGLOSE FINANCIERO", type="primary", use_container_width=True):
        
        df_operacion = df_t1[mask_fechas].copy()
        if fincas_sel:
            df_operacion = df_operacion[df_operacion[col_finca].astype(str).str.upper().str.strip().isin(fincas_sel)]

        if df_operacion.empty:
            st.warning("📭 No se encontraron facturas ejecutadas en ese rango de fechas.")
        else:
            with st.spinner("Desarmando la facturación pieza por pieza..."):
                resultados = []

                for _, row in df_operacion.iterrows():
                    finca_raw = str(row[col_finca]).upper().strip()
                    finca_clean = row['FINCA_CLEAN']
                    coctel = str(row[col_coctel]).upper().strip() if col_coctel else "S/N"
                    ha_num = limpiar_tarifa_excel(row[col_ha])
                    dias_ciclo = row['DIAS_CICLO']

                    # 4. EXTRACCIÓN DE VALORES REALES DE LA TABLA 1
                    costo_vuelo_ha = limpiar_tarifa_excel(row[col_precio_vuelo])
                    costo_vuelo_finca = limpiar_tarifa_excel(row[col_costo_vuelo_finca])
                    costo_x_ha_facturado = limpiar_tarifa_excel(row[col_valor_fact])

                    if ha_num <= 0 or costo_x_ha_facturado <= 0: continue

                    # 5. CRUZAR TARIFA DE SERVICIO TÉCNICO
                    tipo_productor = "TERCERO"
                    if not df_t2.empty:
                        match_t2 = df_t2[df_t2.iloc[:, 0].astype(str).apply(lambda x: re.sub(r'[^A-Z0-9]', '', x.upper().strip())) == finca_clean]
                        if not match_t2.empty and len(match_t2.columns) > 5:
                            tipo_productor = str(match_t2.iloc[0].iloc[5]).strip().upper()

                    if "COOP" in finca_raw or "EMPREBANCOOP" in finca_raw: 
                        tipo_productor = "COOPERATIVA"

                    st_base = 0.0
                    if not df_cfg.empty:
                        match_cfg = df_cfg[df_cfg.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_productor]
                        if not match_cfg.empty and len(match_cfg.columns) > 4:
                            st_base = limpiar_tarifa_excel(match_cfg.iloc[0].iloc[4])

                    if st_base == 0:
                        if tipo_productor == "TERCERO": st_base = 1583.0
                        elif tipo_productor in ["AFILIADO", "COOPERATIVA"]: st_base = 1510.0
                        elif tipo_productor == "ORGANICO": st_base = 1337.0
                        else: st_base = 1337.0

                    # 6. INGENIERÍA INVERSA (El Desglose)
                    resultado_total = math.floor(costo_x_ha_facturado * ha_num)
                    
                    # Costo de Vuelo (usamos directamente la columna V que la tabla ya calculó)
                    if costo_vuelo_finca == 0 and costo_vuelo_ha > 0:
                        costo_vuelo_finca = costo_vuelo_ha * ha_num

                    costo_st_total = math.floor(st_base * dias_ciclo * ha_num)
                    
                    # LO QUE QUEDA ES LA MEZCLA
                    costo_mezcla_total = resultado_total - costo_vuelo_finca - costo_st_total

                    resultados.append({
                        "FINCA": finca_raw,
                        "HECTAREAS": ha_num,
                        "COCTEL": coctel,
                        "DIAS CICLO": dias_ciclo,
                        "PRECIO VUELO": costo_vuelo_ha,
                        "Costo ST ($)": costo_st_total,
                        "Costo Vuelo ($)": costo_vuelo_finca,
                        "Costo Mezcla ($)": costo_mezcla_total,
                        "Costo x Ha ($)": costo_x_ha_facturado,
                        "RESULTADO TOTAL ($)": resultado_total
                    })

                df_resultados = pd.DataFrame(resultados)

                if df_resultados.empty:
                    st.warning("⚠️ No se pudieron generar datos válidos. Verifique que las hectáreas y valores a facturar no sean cero.")
                    return

                # 7. PRESENTACIÓN DE DATOS
                t_st = df_resultados['Costo ST ($)'].sum()
                t_vu = df_resultados['Costo Vuelo ($)'].sum()
                t_mx = df_resultados['Costo Mezcla ($)'].sum()
                t_gr = df_resultados['RESULTADO TOTAL ($)'].sum()

                st.markdown("---")
                st.markdown("### 🎛️ Tablero de Resultados Financieros")
                
                k1, k2, k3, k4 = st.columns(4)
                with k1: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>👨‍🔬 Total Serv. Tec</p><p class='kpi-valor'>$ {formato_latino(t_st, 0)}</p></div>", unsafe_allow_html=True)
                with k2: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>✈️ Total Vuelo</p><p class='kpi-valor'>$ {formato_latino(t_vu, 0)}</p></div>", unsafe_allow_html=True)
                with k3: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>🧪 Total Mezcla Real</p><p class='kpi-valor'>$ {formato_latino(t_mx, 0)}</p></div>", unsafe_allow_html=True)
                with k4: st.markdown(f"<div class='tarjeta-kpi' style='border-left: 5px solid #00ff00;'><p class='kpi-titulo' style='color:#00ff00;'>🔥 FACTURADO (SAP)</p><p class='kpi-valor'>$ {formato_latino(t_gr, 0)}</p></div>", unsafe_allow_html=True)

                df_resumen_finca = df_resultados.groupby('FINCA', as_index=False)[
                    ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']
                ].sum()

                tab1, tab2 = st.tabs(["📊 Detalles Económicos Fila x Fila", "📑 Resumen Ejecutivo por Finca"])
                
                with tab1:
                    df_view = df_resultados.copy()
                    df_view['HECTAREAS'] = df_view['HECTAREAS'].apply(lambda x: formato_latino(x, 2))
                    for col in ["PRECIO VUELO", "Costo ST ($)", "Costo Vuelo ($)", "Costo Mezcla ($)", "Costo x Ha ($)", "RESULTADO TOTAL ($)"]:
                        df_view[col] = df_view[col].apply(lambda x: f"$ {formato_latino(x, 0)}")
                    st.dataframe(df_view, use_container_width=True, hide_index=True)

                with tab2:
                    df_resumen_view = df_resumen_finca.copy()
                    for col in ['Costo ST ($)', 'Costo Vuelo ($)', 'Costo Mezcla ($)', 'RESULTADO TOTAL ($)']:
                        df_resumen_view[col] = df_resumen_view[col].apply(lambda x: f"$ {formato_latino(x, 0)}")
                    st.dataframe(df_resumen_view, use_container_width=True, hide_index=True)

                # 8. EXPORTADOR EXCEL PROFESIONAL
                st.markdown("<br>", unsafe_allow_html=True)
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resultados.to_excel(writer, sheet_name='Detalle_Económico', index=False)
                    df_resumen_finca.to_excel(writer, sheet_name='Resumen_x_Finca', index=False)
                    
                    workbook = writer.book
                    borde = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)

                    for sheet_name in workbook.sheetnames:
                        ws = workbook[sheet_name]
                        ws.sheet_view.showGridLines = False
                        max_r, max_c = ws.max_row, ws.max_column
                        
                        column_headers = {}
                        for col_idx in range(1, max_c + 1):
                            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 20
                            header_val = ws.cell(row=1, column=col_idx).value
                            column_headers[col_idx] = str(header_val).upper() if header_val else ""

                        for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
                            for cell in row:
                                cell.border = borde
                                if cell.row == 1:
                                    cell.fill = header_fill
                                    cell.font = header_font
                                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                else:
                                    cell.alignment = Alignment(vertical='center')
                                    col_name = column_headers.get(cell.column, "")
                                    
                                    if isinstance(cell.value, (int, float)):
                                        if "COSTO" in col_name or "PRECIO" in col_name or "RESULTADO" in col_name or "TOTAL" in col_name:
                                            cell.number_format = '"$" #,##0' 
                                        elif "HECTAREAS" in col_name:
                                            cell.number_format = '#,##0.0'

                st.download_button(
                    label="💾 DESCARGAR AUDITORÍA GERENCIAL (EXCEL)",
                    data=buffer.getvalue(),
                    file_name=f"Auditoria_Facturacion_{fecha_ini}_{fecha_fin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("✅ Ingeniería Inversa completada. La estructura de costos ocultos ha sido revelada.")

if __name__ == "__main__":
    pass
