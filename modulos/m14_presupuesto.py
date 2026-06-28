import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import io

# --- 🔌 CONEXIÓN Y UTILIDADES ---
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except: return None

def a_numero_limpio(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip().replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1:
            partes = v.rsplit('.', 1)
            v = partes[0].replace('.', '') + '.' + partes[1]
        return float(v) if v else 0.0
    except: return 0.0

def parsear_precio(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip()
        v = re.sub(r'[^\d\.,\-]', '', v)
        if not v: return 0.0
        if ',' in v and '.' in v:
            if v.rfind(',') > v.rfind('.'): v = v.replace('.', '').replace(',', '.')
            else: v = v.replace(',', '')
        elif ',' in v:
            if v.count(',') > 1: v = v.replace(',', '')
            else:
                if len(v.split(',')[1]) == 3: v = v.replace(',', '') 
                else: v = v.replace(',', '.')
        elif '.' in v:
            if v.count('.') > 1: v = v.replace('.', '')
            else:
                if len(v.split('.')[1]) == 3: v = v.replace('.', '') 
        return float(v)
    except: return 0.0

def procesar_fecha_pesada(val):
    if pd.isna(val) or str(val).strip() == "": return pd.NaT
    s = str(val).strip()
    if s.replace('.', '', 1).isdigit(): 
        return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y'):
        try: return pd.to_datetime(s, format=fmt)
        except: pass
    try: return pd.to_datetime(s, errors='coerce')
    except: return pd.NaT

def fmt_latino(val, decimales=1):
    try: return f"{float(val):,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(val)

# 💥 PURIFICADOR DE ADITIVOS (Elimina "O", "ONM", etc) 💥
def extraer_receta_rapida(coctel_sel, dict_bases, dict_aditivos_dosis, dict_fertilizantes_dinamico):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if len(partes) > 0 else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = dict_bases.get(base_coctel, {}).copy()

    for aditivo in aditivos:
        nombre_fert = dict_fertilizantes_dinamico.get(aditivo)
        
        if nombre_fert:
            # Si lo encontró exactamente, trae su dosis o asume 0.5
            dosis_fert = dict_aditivos_dosis.get(nombre_fert, 0.5)
            dict_prods[nombre_fert] = dict_prods.get(nombre_fert, 0.0) + dosis_fert
        else:
            # Si el aditivo es un error de digitación ("ONM", "OZN", "O")
            if "NM" in aditivo:
                dict_prods["NATURAMIN WSP"] = dict_prods.get("NATURAMIN WSP", 0.0) + 0.2
            elif "ZN" in aditivo:
                dict_prods["ZINTRAC X LITRO SV"] = dict_prods.get("ZINTRAC X LITRO SV", 0.0) + 0.5
            elif "BT" in aditivo:
                dict_prods["BANATREL SC"] = dict_prods.get("BANATREL SC", 0.0) + 0.5
            # Si no hace match con las siglas salvavidas, SE IGNORA TOTALMENTE. No más químicos fantasma.
    
    if not any("ADHERENTE" in k for k in dict_prods.keys()): dict_prods["ADHERENTE SV"] = 0.13
    if not any("ACONDICIONADOR" in k for k in dict_prods.keys()): 
        dict_prods["ACONDICIONADOR SV"] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
    if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel: 
        dict_prods["IMBIOSIL O"] = 1.5

    return dict_prods

@st.cache_data(show_spinner=False, ttl=7200) 
def descargar_y_masticar_bases():
    gc = inicializar_cliente_gspread()
    if not gc: return pd.DataFrame(), {}, {}, {}, pd.DataFrame(), pd.DataFrame()
    
    boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
    t1_vals = boveda.worksheet("TABLA 1").get_all_values()
    mz_vals = boveda.worksheet("DD_Mesclas").get_all_values()
    cfg_vals = boveda.worksheet("Configuración").get_all_values()
    
    df_t1 = pd.DataFrame(t1_vals[5:], columns=[str(c).upper().strip() for c in t1_vals[4]])
    df_mezclas = pd.DataFrame(mz_vals[1:], columns=[str(c).upper().strip() for c in mz_vals[0]])
    df_cfg = pd.DataFrame(cfg_vals[1:], columns=[str(c).upper().strip() for c in cfg_vals[0]])
    
    col_fecha = next((c for c in df_t1.columns if 'FECHA' in c), 'FECHA')
    col_ha = next((c for c in df_t1.columns if 'NETA' in c or 'FUMIG' in c or 'HECT' in c), None)
    col_coctel = next((c for c in df_t1.columns if 'COCTEL' in c or 'CÓCTEL' in c or 'MEZCLA' in c), None)
    col_pista_t1 = next((c for c in df_t1.columns if 'PISTA' in c or 'BASE' in c), None)

    if col_fecha and col_ha and col_pista_t1 and col_coctel:
        df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
        df_t1 = df_t1.dropna(subset=['FECHA_DT'])
        df_t1['MES'] = df_t1['FECHA_DT'].dt.month
        df_t1['AÑO'] = df_t1['FECHA_DT'].dt.year
        df_t1['HA_CALCULO'] = df_t1[col_ha].apply(a_numero_limpio)
        df_t1['PISTA_OPERATIVA'] = df_t1[col_pista_t1].astype(str).str.upper().str.strip()
        df_t1['COCTEL_NOM'] = df_t1[col_coctel].astype(str).str.upper().str.strip()
    
    dict_bases = {}
    dict_aditivos_dosis = {}
    dict_fert = {}

    if not df_mezclas.empty:
        col_0_limpia = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        for base_name in col_0_limpia.unique():
            if base_name in ["NAN", "", "NONE"]: continue
            rb = df_mezclas[col_0_limpia == base_name]
            prods = {}
            for _, r in rb.iterrows():
                p = str(r.iloc[1]).strip().upper()
                d = a_numero_limpio(r.iloc[2])
                if d > 0 and p not in ['NAN', 'NONE', '']: prods[p] = d
            dict_bases[base_name] = prods
        
        for col_idx in range(len(df_mezclas.columns) - 1):
            for row_idx in range(len(df_mezclas)):
                val_name = str(df_mezclas.iloc[row_idx, col_idx]).strip().upper()
                if val_name and val_name not in ['NAN', 'NONE', '']:
                    val_dosis = a_numero_limpio(df_mezclas.iloc[row_idx, col_idx+1])
                    if val_dosis > 0: dict_aditivos_dosis[val_name] = val_dosis

        if len(df_mezclas.columns) > 13:
            for _, row in df_mezclas.iterrows():
                f_n = str(row.iloc[12]).strip().upper() 
                f_s = str(row.iloc[13]).strip().upper() 
                if f_s and f_n not in ["", "NAN", "NONE", "FERTILIZANTES", "SIGLAS"]:
                    dict_fert[f_s] = f_n

    df_precios_master = pd.DataFrame()
    try:
        sh_precios = gc.open_by_url("https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit")
        precios_consolidados = []
        for ws in sh_precios.worksheets():
            datos_hoja = ws.get_all_values()
            if not datos_hoja: continue
            
            idx_header, col_anio, col_prod, col_precio_tipo = -1, -1, -1, -1
            for i in range(min(10, len(datos_hoja))):
                fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                    idx_header = i
                    col_anio = fila_upper.index('AÑO')
                    col_prod = fila_upper.index('PRODUCTO')
                    col_precio_tipo = next((idx for idx, val in enumerate(fila_upper) if 'PRECIO' in val), -1)
                    break
            
            if idx_header != -1:
                for row in datos_hoja[idx_header+1:]:
                    if col_precio_tipo != -1 and len(row) > col_precio_tipo:
                        if "DOSIS" in str(row[col_precio_tipo]).upper(): continue
                            
                    if len(row) > max(col_anio, col_prod):
                        anio_str = str(row[col_anio]).strip()
                        str_prod = str(row[col_prod]).strip().upper()
                        if anio_str.isdigit() and str_prod:
                            col_inicio = max(col_anio, col_prod) + 1
                            vals = [parsear_precio(v) for v in row[col_inicio:] if str(v).strip() != ""]
                            vals = [v for v in vals if v > 0]
                            prom = sum(vals)/len(vals) if vals else 0.0
                            if prom > 0:
                                precios_consolidados.append({
                                    'AÑO': int(anio_str), 
                                    'PRODUCTO': str_prod, 
                                    'PROD_CLEAN': re.sub(r'[^\w]', '', str_prod), 
                                    'PRECIO': prom
                                })
        df_precios_master = pd.DataFrame(precios_consolidados)
    except: pass
        
    return df_t1, dict_bases, dict_aditivos_dosis, dict_fert, df_cfg, df_precios_master

def extraer_precios_maestros(df_cfg):
    precios = {}
    if df_cfg.empty: return precios
    c_p_i, c_c_i = 8, 9
    for i in range(min(5, len(df_cfg))):
        r_c = [str(x).upper().strip() for x in df_cfg.iloc[i].tolist()]
        if 'PRODUCTO' in r_c and 'COSTO' in r_c:
            c_p_i, c_c_i = r_c.index('PRODUCTO'), r_c.index('COSTO')
            break
    for r in range(len(df_cfg)):
        p = str(df_cfg.iloc[r, c_p_i]).upper().strip()
        c = parsear_precio(df_cfg.iloc[r, c_c_i])
        if p and p not in ["NAN", "NONE", ""]: precios[p] = c
    return precios

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-presupuesto { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    div[data-testid="stDataFrame"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    .kpi-presupuesto { background-color: #0d1b2a; color: white; padding: 20px; border-radius: 10px; border-left: 6px solid #d4af37; box-shadow: 0 4px 6px rgba(0,0,0,0.2); margin-bottom: 15px;}
    .kpi-titulo { color: #d4af37; font-weight: bold; font-size: 14px; margin-bottom: 5px; text-transform: uppercase; }
    .kpi-valor { font-size: 32px; font-weight: 900; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-presupuesto'>💰 Módulo 14: Pronóstico Financiero</h1>", unsafe_allow_html=True)
    st.write("Proyección estratégica del flujo de efectivo con rastreo de precios históricos y ajuste de inflación.")

    st.markdown("### ⚙️ Parámetros del Presupuesto")
    
    fila1_col1, fila1_col2, fila1_col3 = st.columns(3)
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    opciones_mes = ["📊 AÑO COMPLETO (TODOS)"] + list(meses_dict.values())
    
    mes_sel = fila1_col1.selectbox("📅 Mes a Proyectar:", opciones_mes)
    pista_sel = fila1_col2.selectbox("📍 Base Operativa:", ["TODAS", "PLUC", "PORI", "PDIV", "TEHO", "LUCI"])
    anio_actual = datetime.now().year
    anio_presupuesto = fila1_col3.selectbox("🎯 Año a Presupuestar:", [2026, 2027, 2028, 2029, 2030], index=1)
    
    fila2_col1, fila2_col2, fila2_col3 = st.columns(3)
    profundidad_sel = fila2_col1.selectbox("🔍 Base Histórica de Consumo:", ["Último Año", "Últimos 2 Años", "Últimos 3 Años", "Histórico Completo"])
    crecimiento_sel = fila2_col2.number_input("📈 Crecimiento Operativo (%)", min_value=-50, max_value=200, value=0, step=5)
    inflacion_sel = fila2_col3.number_input("💸 Inflación Anual Estimada (%)", min_value=0.0, max_value=50.0, value=8.0, step=1.0)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 CALCULAR PRESUPUESTO FINANCIERO", type="primary", use_container_width=True):
        with st.spinner("Compilando matriz a velocidad de memoria nativa..."):
            try:
                df_t1_base, dict_bases, dict_aditivos_dosis, dict_fert, df_cfg, df_precios_master = descargar_y_masticar_bases()
                
                if df_t1_base.empty:
                    st.error("🚨 No se pudo establecer conexión con las Bóvedas de Datos.")
                    return

                df_t1 = df_t1_base.copy()
                dict_precios_backup = extraer_precios_maestros(df_cfg)
                año_base = 2026 

                if profundidad_sel == "Último Año": df_t1 = df_t1[df_t1['AÑO'] >= (anio_actual - 1)]
                elif profundidad_sel == "Últimos 2 Años": df_t1 = df_t1[df_t1['AÑO'] >= (anio_actual - 2)]
                elif profundidad_sel == "Últimos 3 Años": df_t1 = df_t1[df_t1['AÑO'] >= (anio_actual - 3)]

                if mes_sel != "📊 AÑO COMPLETO (TODOS)":
                    mes_num = next(k for k, v in meses_dict.items() if v == mes_sel)
                    df_t1 = df_t1[df_t1['MES'] == mes_num]
                
                total_anios_boveda = df_t1['AÑO'].nunique()
                if total_anios_boveda == 0: total_anios_boveda = 1

                traductor_pistas = {"PLUC": "FUMIGARAY", "PORI": "AEROPENOR", "LUCI": "GENESYS", "TEHO": "AVIL", "PDIV": "ASA"}
                if pista_sel != "TODAS":
                    pista_t1_esperada = traductor_pistas.get(pista_sel, pista_sel)
                    df_t1 = df_t1[df_t1['PISTA_OPERATIVA'].str.contains(pista_t1_esperada, na=False)]

                consumo_esperado = {} 
                ha_total_detectada = 0.0

                if not df_t1.empty:
                    ha_total_detectada = df_t1['HA_CALCULO'].sum()
                    ha_total_por_coctel = df_t1.groupby(['PISTA_OPERATIVA', 'COCTEL_NOM'])['HA_CALCULO'].sum().reset_index()
                    factor_crecimiento = 1 + (crecimiento_sel / 100.0)
                    ha_total_por_coctel['HA_PROYECTADA'] = (ha_total_por_coctel['HA_CALCULO'] / total_anios_boveda) * factor_crecimiento

                    for _, row_c in ha_total_por_coctel.iterrows():
                        coctel_completo = str(row_c['COCTEL_NOM'])
                        ha_proyectada = row_c['HA_PROYECTADA']

                        receta_dict = extraer_receta_rapida(coctel_completo, dict_bases, dict_aditivos_dosis, dict_fert)
                        for prod_quimico, dosis in receta_dict.items():
                            consumo_esperado[prod_quimico] = consumo_esperado.get(prod_quimico, 0) + (dosis * ha_proyectada)

                resultados = []
                gran_total_presupuesto = 0.0
                precios_records = df_precios_master.to_dict('records') if not df_precios_master.empty else []

                for producto, volumen in consumo_esperado.items():
                    if volumen > 0:
                        precio_unitario_final = 0.0
                        precio_hist_base = 0.0
                        anio_origen = ""
                        p_clean = re.sub(r'[^\w]', '', producto.upper().strip())
                        
                        for r_db in precios_records:
                            if r_db['AÑO'] == año_base:
                                if p_clean in r_db['PROD_CLEAN'] or r_db['PROD_CLEAN'] in p_clean:
                                    precio_hist_base = r_db['PRECIO']
                                    anios_pasados = max(0, anio_presupuesto - año_base)
                                    precio_unitario_final = precio_hist_base * ((1 + (inflacion_sel / 100.0)) ** anios_pasados)
                                    anio_origen = f"Base {año_base} (+{inflacion_sel}% x {anios_pasados}a)"
                                    break
                        
                        if precio_unitario_final == 0.0 and precios_records:
                            matches_hist = []
                            for r_db in precios_records:
                                if r_db['AÑO'] < año_base:
                                    if p_clean in r_db['PROD_CLEAN'] or r_db['PROD_CLEAN'] in p_clean:
                                        matches_hist.append(r_db)
                            
                            if matches_hist:
                                best_match = max(matches_hist, key=lambda x: x['AÑO'])
                                anio_hist = int(best_match['AÑO'])
                                precio_hist_base = best_match['PRECIO']
                                
                                anios_pasados = max(0, anio_presupuesto - anio_hist)
                                precio_unitario_final = precio_hist_base * ((1 + (inflacion_sel / 100.0)) ** anios_pasados)
                                anio_origen = f"Rescatado {anio_hist} (+{inflacion_sel}% x {anios_pasados}a)"

                        if precio_unitario_final == 0.0:
                            precio_bk = dict_precios_backup.get(producto, 0.0)
                            if precio_bk < 1000:
                                for p_bk, val_bk in dict_precios_backup.items():
                                    bk_clean = re.sub(r'[^\w]', '', p_bk.upper().strip())
                                    if p_clean in bk_clean or bk_clean in p_clean:
                                        if val_bk >= 1000: 
                                            precio_hist_base = val_bk
                                            break
                            else:
                                precio_hist_base = precio_bk
                                
                            if precio_hist_base >= 1000:
                                anios_pasados = max(0, anio_presupuesto - anio_actual)
                                precio_unitario_final = precio_hist_base * ((1 + (inflacion_sel / 100.0)) ** anios_pasados)
                                anio_origen = f"Conf. Local (+{inflacion_sel}% x {anios_pasados}a)"

                        if precio_unitario_final == 0.0:
                            anio_origen = "⚠️ Falta Precio"

                        presupuesto_prod = volumen * precio_unitario_final
                        gran_total_presupuesto += presupuesto_prod
                        
                        resultados.append({
                            "🧪 INSUMO QUÍMICO": producto,
                            "📦 VOLUMEN ESTIMADO": volumen,
                            "💵 PRECIO BASE": precio_hist_base,
                            "📈 PRECIO AJUSTADO": precio_unitario_final,
                            "🔎 ORIGEN": anio_origen,
                            "💰 PRESUPUESTO TOTAL": presupuesto_prod
                        })

                df_presupuesto = pd.DataFrame(resultados)

                st.markdown("---")
                
                if df_presupuesto.empty:
                    st.warning("⚠️ No hay suficientes datos históricos para proyectar este escenario.")
                else:
                    st.markdown(f"""
                    <div class='kpi-presupuesto'>
                        <div class='kpi-titulo'>FLUJO DE EFECTIVO PROYECTADO PARA {anio_presupuesto} ({mes_sel})</div>
                        <p class='kpi-valor'>$ {fmt_latino(gran_total_presupuesto, 0)} <span style='font-size: 16px; font-weight: normal;'>COP</span></p>
                        <p style='margin: 0; margin-top: 10px; color: #a0aec0; font-size: 12px;'>Calculado sobre {fmt_latino(ha_total_detectada/total_anios_boveda * (1 + crecimiento_sel/100))} Hectáreas proyectadas. Factor Crecimiento: {crecimiento_sel}% | Inflación de ajuste: {inflacion_sel}%</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    df_presupuesto = df_presupuesto.sort_values(by="🧪 INSUMO QUÍMICO", ascending=True)
                    df_vista = df_presupuesto.copy()
                    
                    df_vista['📦 VOLUMEN ESTIMADO'] = df_vista['📦 VOLUMEN ESTIMADO'].apply(lambda x: fmt_latino(x, 1))
                    
                    def formatear_precio(x): return f"$ {fmt_latino(x, 0)}" if x > 0 else "---"
                    
                    df_vista['💵 PRECIO BASE'] = df_vista['💵 PRECIO BASE'].apply(formatear_precio)
                    df_vista['📈 PRECIO AJUSTADO'] = df_vista['📈 PRECIO AJUSTADO'].apply(formatear_precio)
                    df_vista['💰 PRESUPUESTO TOTAL'] = df_vista['💰 PRESUPUESTO TOTAL'].apply(lambda x: f"$ {fmt_latino(x, 0)}" if x > 0 else "$ 0")

                    st.markdown("### 📋 Desglose Financiero por Insumo")
                    
                    def color_origen(val):
                        if "(+" in str(val) and "x 0a" not in str(val): return 'color: #d4af37;'
                        if "⚠️" in str(val): return 'color: #cc0000; font-weight: bold;'
                        return 'color: #155724;'

                    styled_df = df_vista.style.map(color_origen, subset=['🔎 ORIGEN']).set_properties(**{'text-align': 'left'})
                    
                    st.dataframe(
                        styled_df, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "🧪 INSUMO QUÍMICO": st.column_config.TextColumn("INSUMO QUÍMICO", width="medium"),
                            "📦 VOLUMEN ESTIMADO": st.column_config.TextColumn("VOLUMEN", width="small"),
                            "💵 PRECIO BASE": st.column_config.TextColumn("PRECIO BASE", width="small"),
                            "📈 PRECIO AJUSTADO": st.column_config.TextColumn("PRECIO AJUSTADO", width="small"),
                            "🔎 ORIGEN": st.column_config.TextColumn("ORIGEN DEL DATO", width="medium"),
                            "💰 PRESUPUESTO TOTAL": st.column_config.TextColumn("PRESUPUESTO TOTAL", width="medium")
                        }
                    )

                    st.markdown("<br>", unsafe_allow_html=True)
                    col_down1, col_down2, col_down_vacia = st.columns([1, 1, 2])
                    
                    try:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer) as writer:
                            df_vista.to_excel(writer, sheet_name='Presupuesto', index=False)
                        
                        col_down1.download_button(
                            label="📊 Exportar a Excel (Recomendado)",
                            data=buffer.getvalue(),
                            file_name=f"Presupuesto_{anio_presupuesto}_{mes_sel}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except:
                        col_down1.error("⚠️ Instale 'openpyxl' en su entorno para habilitar Excel.")

                    csv_data = df_vista.to_csv(index=False).encode('utf-8')
                    col_down2.download_button(
                        label="📄 Exportar a CSV",
                        data=csv_data,
                        file_name=f"Presupuesto_{anio_presupuesto}_{mes_sel}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"🚨 Falla en los cálculos financieros: {e}")
