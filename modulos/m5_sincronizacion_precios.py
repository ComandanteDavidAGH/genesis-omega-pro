import streamlit as st
import pandas as pd
import numpy as np
import gspread
import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

@st.cache_resource(show_spinner=False)
def obtener_cliente_gspread_unificado():
    """ Centraliza la autenticación unificada con Google Cloud una sola vez en RAM """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    if "gcp_credentials" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_credentials"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    try:
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

def purificar_y_convertir_precio(valor_crudo):
    if not valor_crudo:
        return 0.0
    val_str = str(valor_crudo).replace("$", "").replace("COP", "").replace(" ", "").strip()
    if "," in val_str and "." in val_str:
        if val_str.find(".") < val_str.find(","):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        if len(val_str.split(",")[-1]) == 2:
            val_str = val_str.replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    try:
        return float(val_str)
    except ValueError:
        return 0.0

# =================================================================
# 📊 EXPORTADOR EXCEL PREMIUM (GERENCIA)
# =================================================================
def generar_excel_gerencial(df_comp):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_comp.to_excel(writer, sheet_name='Rentabilidad_Gerencial', index=False)
        ws = writer.sheets['Rentabilidad_Gerencial']

        header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        header_font = Font(color="D4AF37", bold=True, size=11)
        borde_fino = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
                            top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
        
        for col_num, col_name in enumerate(df_comp.columns, 1):
            col_letter = openpyxl.utils.get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = 22
            
            cell_header = ws.cell(row=1, column=col_num)
            cell_header.fill = header_fill
            cell_header.font = header_font
            cell_header.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell_header.border = borde_fino
            
            for row_num in range(2, len(df_comp) + 2):
                cell = ws.cell(row=row_num, column=col_num)
                cell.border = borde_fino
                cell.alignment = Alignment(vertical='center')
                if "HA" in col_name and "DOSIS" not in col_name and "%" not in col_name:
                    cell.number_format = '"$" #,##0'
                elif "%" in col_name:
                    cell.number_format = '0.00 "%"'
                elif "DOSIS" in col_name:
                    cell.number_format = '#,##0.00'
                    
    return buffer.getvalue()

# =================================================================
# 🔌 MEMORIA DINÁMICA DE MÁRGENES Y TARIFAS
# =================================================================
@st.cache_data(show_spinner=False, ttl=1800)
def obtener_tarifario_maestro_cached(_supabase_client):
    df = pd.DataFrame()
    if _supabase_client:
        try:
            respuesta = _supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
            if respuesta.data: df = pd.DataFrame(respuesta.data)
        except Exception: df = pd.DataFrame()

    margenes = {"TERCERO": 1.451, "AFILIADO": 1.164, "COOPERATIVA / SOCIO": 1.112, "ORGANICO": 1.011}
    
    gc = obtener_cliente_gspread_unificado()
    if gc:
        try:
            sh = gc.open_by_url(URL_BOVEDA_MAESTRA)
            ws = sh.worksheet("Configuración")
            datos = ws.get_all_values()
            
            if len(datos) > 1:
                df_raw = pd.DataFrame(datos[1:], columns=datos[0])
                def parse_mult(v_str):
                    try:
                        v = str(v_str).strip().replace(",", ".")
                        if "%" in v: return 1 + (float(v.replace("%", "")) / 100.0)
                        vf = float(v)
                        if 1.0 <= vf <= 2.0: return vf
                        if 1000 <= vf <= 2000: return vf / 1000.0 
                    except: pass
                    return 0.0

                for idx, row in df_raw.iterrows():
                    grupo = str(row.iloc[0]).strip().upper()
                    if grupo in ["TERCERO", "AFILIADO", "SOCIO", "COOPERATIVA", "ORGANICO"]:
                        factor = 0.0
                        for c in range(1, 4):
                            f_val = parse_mult(row.iloc[c])
                            if f_val > 1.0:
                                factor = f_val
                                break
                        if factor > 0:
                            if grupo in ["SOCIO", "COOPERATIVA"]: margenes["COOPERATIVA / SOCIO"] = factor
                            elif grupo == "TERCERO": margenes["TERCERO"] = factor
                            elif grupo == "AFILIADO": margenes["AFILIADO"] = factor
                            elif grupo == "ORGANICO": margenes["ORGANICO"] = factor
                
                if df.empty and len(df_raw.columns) > 10:
                    df = df_raw.iloc[:, [8, 10]].copy()
                    df.columns = ['PRODUCTO', 'COSTO']
        except Exception: pass

    if df.empty or 'PRODUCTO' not in df.columns or 'COSTO' not in df.columns:
        return pd.DataFrame(), []
        
    df['PRODUCTO'] = df['PRODUCTO'].astype(str).str.strip().str.upper()
    mask_validos = (df['PRODUCTO'].notna() & (df['PRODUCTO'] != "") & (df['PRODUCTO'] != "PRODUCTO") & (~df['PRODUCTO'].str.contains("INVENTARIO", na=False)))
    df = df[mask_validos].copy()
    
    df['COSTO BASE'] = df['COSTO'].apply(purificar_y_convertir_precio)
    df = df[df['COSTO BASE'] > 0].copy()
    
    if df.empty: return pd.DataFrame(), []
    
    def fmt_pct(factor): return round((factor - 1.0) * 100, 2)
    
    col_ter = f"TERCERO (+{fmt_pct(margenes['TERCERO'])}%)"
    col_afi = f"AFILIADO (+{fmt_pct(margenes['AFILIADO'])}%)"
    col_soc = f"COOP/SOCIO (+{fmt_pct(margenes['COOPERATIVA / SOCIO'])}%)"
    col_org = f"ORGÁNICO (+{fmt_pct(margenes['ORGANICO'])}%)"

    df[col_ter] = (df['COSTO BASE'] * margenes['TERCERO']).round(0)
    df[col_afi] = (df['COSTO BASE'] * margenes['AFILIADO']).round(0)
    df[col_soc] = (df['COSTO BASE'] * margenes['COOPERATIVA / SOCIO']).round(0)
    df[col_org] = (df['COSTO BASE'] * margenes['ORGANICO']).round(0)
    
    cols = ["PRODUCTO", "COSTO BASE", col_ter, col_afi, col_soc, col_org]
    df_tarifario = df[cols].sort_values(by="PRODUCTO").reset_index(drop=True)
    
    return df_tarifario, cols[1:]

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE TARIFAS Y MACRO OMEGA V12
# =================================================================

def ejecutar(supabase_client, extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    .hud-tarifas { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }
    .hud-tarifas-item { text-align: center; flex: 1; }
    .hud-tarifas-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-tarifas-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    
    div[data-testid="stTextInput"] input, 
    div[data-testid="stNumberInput"] input, 
    div[data-testid="stSelectbox"] > div { 
        border: 2px solid #0d1b2a !important; 
        border-radius: 6px !important; 
        background-color: #ffffff !important;
        color: #0d1b2a !important;
        font-weight: 800 !important; 
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    gc = obtener_cliente_gspread_unificado()

    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Cómputo de Perfiles)")
        
        col_t1, col_t2 = st.columns([3, 1])
        with col_t2:
            if st.button("🔄 RECARGAR TARIFARIO", use_container_width=True, type="secondary"):
                st.cache_data.clear()
                st.session_state.pop('df_tarifario', None)
                st.session_state.pop('opciones_cols_m5', None)
                st.rerun()

        if 'df_tarifario' not in st.session_state or st.session_state['df_tarifario'].empty:
            df_tarifario_cached, opciones_cols = obtener_tarifario_maestro_cached(supabase_client)
            if not df_tarifario_cached.empty:
                st.session_state['df_tarifario'] = df_tarifario_cached
                st.session_state['opciones_cols_m5'] = opciones_cols

        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_t = st.session_state['df_tarifario']
            cols_dinamicas = st.session_state.get('opciones_cols_m5', [])
            
            if not cols_dinamicas or len(cols_dinamicas) < 2:
                df_t, cols_dinamicas = obtener_tarifario_maestro_cached(supabase_client)
                st.session_state['df_tarifario'] = df_t
                st.session_state['opciones_cols_m5'] = cols_dinamicas
            
            total_quimicos_tarifados = len(df_t)
            costo_maximo_comercial = df_t[cols_dinamicas[1]].max() if len(cols_dinamicas) > 1 else 0
            costo_medio_base = df_t['COSTO BASE'].mean() if 'COSTO BASE' in df_t.columns else 0
            
            st.markdown(f"""
            <div class="hud-tarifas">
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Insumos en Línea</p>
                    <p class="hud-tarifas-value">🧪 {total_quimicos_tarifados} ítems</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Costo Promedio</p>
                    <p class="hud-tarifas-value">💵 $ {costo_medio_base:,.0f}</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Tope Máximo</p>
                    <p class="hud-tarifas-value">📈 $ {costo_maximo_comercial:,.0f}</p>
                </div>
            </div>
            """.replace(",", "."), unsafe_allow_html=True)
            
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Comparativo Gerencial (Utilidad/Ha)"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes Oficiales")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### 📟 Consola de Copiado Masivo para SAP")
                c_cop1, c_cop2 = st.columns([2, 1])
                with c_cop1:
                    lista_opciones = list(reversed(cols_dinamicas)) if cols_dinamicas else ["COSTO BASE"]
                    col_margen = st.selectbox("Seleccione el Perfil de Productor:", lista_opciones, key="sb_perfil_copia")
                with c_cop2:
                    st.write("") 
                    st.write("")
                    incluir_nombres = st.toggle("🏷️ Incluir Nombres de Productos", value=False, key="toggle_inc_nombres")

                if incluir_nombres:
                    max_len = max([len(str(p)) for p in df_t["PRODUCTO"]] + [35]) + 4
                    lista_textos = [f"{str(p).ljust(max_len)}│  {fmt_sap(v)}" for p, v in zip(df_t["PRODUCTO"], df_t[col_margen])]
                else:
                    lista_textos = [fmt_sap(x) for x in df_t[col_margen]]
                
                st.info(f"📊 **Listos para Transferencia:** {len(lista_textos)} registros procesados en columna alineada. Haga clic en 📋 para copiar.")
                st.code("\n".join(lista_textos), language="text")
                    
            with t3:
                st.markdown("#### 🎯 Análisis Gerencial: Rentabilidad y Margen por Hectárea")
                st.info("💡 **SIMULADOR DE UTILIDAD:** Seleccione los insumos y defina la Dosis exacta por Hectárea. El sistema calculará la utilidad neta y el % de rentabilidad para cada perfil comercial cruzado.")
                
                opciones_productos = df_t["PRODUCTO"].tolist()
                prods_sel = st.multiselect(
                    "🔍 Seleccione los Productos a Analizar:", 
                    options=opciones_productos,
                    default=[opciones_productos[0]] if opciones_productos else []
                )
                
                if prods_sel:
                    st.markdown("##### ⚖️ 1. Definir Dosis (L/Kg por Hectárea)")
                    
                    df_dosis_base = pd.DataFrame({"PRODUCTO": prods_sel, "DOSIS (L/Kg/Ha)": [1.0] * len(prods_sel)})
                    
                    df_dosis = st.data_editor(
                        df_dosis_base, 
                        hide_index=True, 
                        use_container_width=True,
                        column_config={
                            "PRODUCTO": st.column_config.TextColumn("🧪 Producto", disabled=True),
                            "DOSIS (L/Kg/Ha)": st.column_config.NumberColumn("⚖️ Dosis/Ha (Doble clic para editar)", min_value=0.001, format="%.2f", step=0.1)
                        }
                    )
                    
                    dosis_dict = dict(zip(df_dosis["PRODUCTO"], df_dosis["DOSIS (L/Kg/Ha)"]))
                    
                    st.markdown("##### 📊 2. Matriz de Rentabilidad Financiera")
                    
                    matriz_gerencial = []
                    for prod_sel in prods_sel:
                        datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
                        dosis = dosis_dict.get(prod_sel, 1.0)
                        costo_base_ha = datos_prod["COSTO BASE"] * dosis
                        
                        for col_margen in cols_dinamicas[1:]:
                            precio_venta_ha = datos_prod[col_margen] * dosis
                            utilidad_dinero = precio_venta_ha - costo_base_ha
                            utilidad_pct = (utilidad_dinero / costo_base_ha * 100) if costo_base_ha > 0 else 0
                            
                            perfil = str(col_margen).split("(+")[0].strip()
                            
                            matriz_gerencial.append({
                                "🧪 PRODUCTO": prod_sel,
                                "⚖️ DOSIS/HA": dosis,
                                "🤝 PERFIL COMERCIAL": perfil,
                                "📉 COSTO BASE/HA": costo_base_ha,
                                "📈 PRECIO VENTA/HA": precio_venta_ha,
                                "💰 UTILIDAD NETA/HA": utilidad_dinero,
                                "🚀 MARGEN (%)": utilidad_pct
                            })
                    
                    df_gerencial = pd.DataFrame(matriz_gerencial)
                    
                    def colorear_utilidad(val):
                        if val > 30: return 'color: #27AE60; font-weight: bold;'
                        if val > 15: return 'color: #2980B9; font-weight: bold;'
                        if val > 0: return 'color: #F1C40F; font-weight: bold;'
                        return 'color: #E74C3C; font-weight: bold;'
                        
                    st.dataframe(
                        df_gerencial.style.map(lambda x: 'color: #27AE60; font-weight: bold;' if x > 0 else '', subset=['💰 UTILIDAD NETA/HA'])
                                        .map(colorear_utilidad, subset=['🚀 MARGEN (%)']),
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "🧪 PRODUCTO": st.column_config.TextColumn("🧪 PRODUCTO"),
                            "⚖️ DOSIS/HA": st.column_config.NumberColumn("⚖️ DOSIS/HA", format="%.2f"),
                            "🤝 PERFIL COMERCIAL": st.column_config.TextColumn("🤝 PERFIL COMERCIAL"),
                            "📉 COSTO BASE/HA": st.column_config.NumberColumn("📉 COSTO BASE/HA", format="$ %d"),
                            "📈 PRECIO VENTA/HA": st.column_config.NumberColumn("📈 PRECIO VENTA/HA", format="$ %d"),
                            "💰 UTILIDAD NETA/HA": st.column_config.NumberColumn("💰 UTILIDAD NETA/HA", format="$ %d"),
                            "🚀 MARGEN (%)": st.column_config.NumberColumn("🚀 MARGEN (%)", format="%.2f %%"),
                        }
                    )
                    
                    excel_data = generar_excel_gerencial(df_gerencial)
                    st.download_button(
                        label="📥 DESCARGAR REPORTE GERENCIAL (EXCEL VIP)", 
                        data=excel_data, 
                        file_name=f"Reporte_Rentabilidad_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                        use_container_width=True,
                        type="primary"
                    )
        else:
            st.warning("⚠️ No se detectaron datos en el tarifario. Haga clic en 'Recargar Tarifario' para sincronizar con la nube.")

    # --- 🚀 SECCIÓN INFERIOR: OMEGA V12 ---
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    
    with st.container(border=True):
        c_url1, c_url2 = st.columns(2)
        with c_url1:
            st.text_input("🔗 1. Base de Origen Activa:", value="DATABASE: Supabase Cloud / Drive Fallback [PRECIOS_INSUMOS]", disabled=True)
        with c_url2:
            url_dest = st.text_input("🎯 2. URL de Sábana Destino (Google Sheets):", placeholder="Pegue el enlace completo aquí...")
        
        semana_target = st.number_input("🔢 Digite la Semana a actualizar (1 a 53):", min_value=1, max_value=53, value=24, step=1)
        
        c_btn1, c_btn2 = st.columns(2)
        
        dict_precios = {}
        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_m = st.session_state['df_tarifario']
            dict_precios = dict(zip(df_m['PRODUCTO'], df_m['COSTO BASE']))

        with c_btn1:
            if st.button("📊 PREVISUALIZAR COMPORTAMIENTO DE PRECIOS POR DOSIS", use_container_width=True, type="secondary"):
                if gc is None:
                    st.error("🚨 Enlace satelital roto con Google Cloud.")
                elif not url_dest or "http" not in url_dest:
                    st.error("❌ Ingrese la URL de destino para previsualizar.")
                elif not dict_precios:
                    st.error("⚠️ El tarifario no está cargado. Cargue el tarifario antes de previsualizar.")
                else:
                    try:
                        with st.spinner("🕵️‍♂️ Calculando Comportamiento Operativo desde la RAM..."):
                            sh_dest = gc.open_by_url(url_dest)
                            ws_datos = sh_dest.worksheet("DATOS")
                            datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                            
                            idx_fila_semanas = 6
                            for idx, r in enumerate(datos_dest[:12]):
                                r_str = [str(cell).strip().split('.')[0] for cell in r]
                                if any(w in r_str for w in ["11", "12", "13", "18"]):
                                    idx_fila_semanas = idx
                                    break
                            
                            filas_comp = []
                            for r_idx, row in enumerate(datos_dest):
                                n_fila = r_idx + 1
                                if n_fila < (idx_fila_semanas + 2): continue
                                
                                row_padded = row + [""] * (15 - len(row)) if len(row) < 15 else row
                                tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                                producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                                
                                if producto_dest in dict_precios:
                                    precio_pleno = dict_precios[producto_dest]
                                    dosis_valor = extraer_numero(row_padded[0])
                                    
                                    if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                        valor_dosis = precio_pleno * dosis_valor if dosis_valor > 0 else 0
                                        formula = f"{dosis_valor} Dosis × $ {precio_pleno:,.0f}"
                                    else:
                                        valor_dosis = precio_pleno
                                        formula = "Precio Unitario Directo"
                                        
                                    filas_comp.append({
                                        "Fila SAP": n_fila,
                                        "Tipo de Registro": tipo_tabla,
                                        "Insumo / Producto": producto_dest,
                                        "Precio Final Calculado": valor_dosis,
                                        "Lógica de Impacto": formula
                                    })
                            
                            if filas_comp:
                                df_vis = pd.DataFrame(filas_comp).copy()
                                df_vis["Precio Final Calculado"] = df_vis["Precio Final Calculado"].map("$ {:,.0f}".format).str.replace(",", ".")
                                st.dataframe(df_vis, use_container_width=True, hide_index=True)
                            else:
                                st.warning("⚠️ No se encontraron coincidencias de insumos en la hoja DATOS.")
                    except Exception as e:
                        st.error(f"🚨 Falla en análisis: {e}")

        with c_btn2:
            if st.button("🚀 EJECUTAR SINCRONIZACIÓN OMEGA V12", use_container_width=True, type="primary"):
                if gc is None:
                    st.error("🚨 Enlace satelital roto con Google Cloud.")
                    return
                elif not url_dest or "http" not in url_dest:
                    st.error("❌ Digite una URL válida de destino.")
                    return
                elif not dict_precios:
                    st.error("⚠️ Tarifario no disponible en memoria.")
                    return
                try:
                    with st.status("🕵️‍♂️ CONECTANDO CON CÉLULA DE PRECIOS Y DESTINO...", expanded=True) as status:
                        sh_dest = gc.open_by_url(url_dest)
                        ws_datos = sh_dest.worksheet("DATOS")
                        datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                        
                        idx_fila_semanas = 6
                        for idx, r in enumerate(datos_dest[:12]):
                            r_str = [str(cell).strip().split('.')[0] for cell in r]
                            if any(w in r_str for w in ["11", "12", "13", "18"]):
                                idx_fila_semanas = idx
                                break
                        
                        col_semana = -1
                        for i, v in enumerate(datos_dest[idx_fila_semanas]):
                            if str(v).strip().split('.')[0] == str(semana_target):
                                col_semana = i + 1
                                break
                        
                        if col_semana == -1: col_semana = int(semana_target) + 5
                        
                        updates = [{'range': gspread.utils.rowcol_to_a1(idx_fila_semanas + 1, col_semana), 'values': [[int(semana_target)]]}]
                        
                        for r_idx, row in enumerate(datos_dest):
                            n_fila = r_idx + 1
                            if n_fila < (idx_fila_semanas + 2): continue
                            
                            row_padded = row + [""] * (max(col_semana + 2, 15) - len(row)) if len(row) < max(col_semana + 2, 15) else row
                            tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                            producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                            
                            if producto_dest in dict_precios:
                                precio_unitario = dict_precios[producto_dest]
                                if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                    dosis_valor = extraer_numero(row_padded[0])
                                    valor_final = precio_unitario * dosis_valor if dosis_valor > 0 else 0
                                else:
                                    valor_final = precio_unitario
                                    
                                updates.append({'range': gspread.utils.rowcol_to_a1(n_fila, col_semana), 'values': [[valor_final]]})

                        if len(updates) > 1:
                            ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                            status.update(label="🎯 ¡MÓDULO DE DOSIS AJUSTADO!", state="complete")
                            st.success(f"🎉 Precios impactados en la columna {col_semana} de la macro para la semana {semana_target}.")
                            st.balloons()
                        else:
                            st.warning("⚠️ No se generaron actualizaciones de precios.")
                except Exception as e:
                    st.error(f"🚨 FALLA EN LA INYECCIÓN: {e}")

if __name__ == "__main__":
    pass
