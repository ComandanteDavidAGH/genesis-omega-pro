import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
from oauth2client.service_account import ServiceAccountCredentials

# --- 🔌 CONEXIÓN Y UTILIDADES ---
@st.cache_resource(show_spinner=False)
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

# 💥 PARSEADOR FINANCIERO LATINO COP 💥
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

def obtener_dosis_fertilizante(df_mezclas, fert_name):
    try:
        for col_idx in range(len(df_mezclas.columns) - 1):
            mask = df_mezclas.iloc[:, col_idx].astype(str).str.strip().str.upper() == fert_name
            if mask.any():
                val = pd.to_numeric(df_mezclas[mask].iloc[0, col_idx+1], errors='coerce')
                if pd.notna(val) and val > 0: return float(val)
    except: pass
    return None 

def extraer_receta_completa(coctel_sel, df_mezclas, dict_fertilizantes_dinamico):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if len(partes) > 0 else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    if not df_mezclas.empty:
        col_0_limpia = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0_limpia == base_coctel]
        for _, r in rb.iterrows():
            p = str(r.iloc[1]).strip().upper()
            d = a_numero_limpio(r.iloc[2])
            if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    for aditivo in aditivos:
        if aditivo in dict_fertilizantes_dinamico:
            nombre_fert = dict_fertilizantes_dinamico[aditivo]
            dosis_fert = obtener_dosis_fertilizante(df_mezclas, nombre_fert)
            if dosis_fert is None:
                if "NATURAMIN" in nombre_fert: dosis_fert = 0.2
                elif "ZINTRAC" in nombre_fert: dosis_fert = 0.5
                elif "BANATREL" in nombre_fert: dosis_fert = 0.5
                else: dosis_fert = 0.5
            dict_prods[nombre_fert] = dict_prods.get(nombre_fert, 0.0) + dosis_fert
    
    if not any("ADHERENTE" in k for k in dict_prods.keys()): dict_prods["ADHERENTE SV"] = 0.13
    if not any("ACONDICIONADOR" in k for k in dict_prods.keys()): 
        dict_prods["ACONDICIONADOR SV"] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
    if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel: 
        dict_prods["IMBIOSIL O"] = 1.5

    return dict_prods

# 🧠 BACKUP RAT LOGÍSTICO: Solo por si falta en toda la sábana de precios
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

# 🧠 RECOLECTOR MAESTRO DE PRECIOS (DATOS E HISTORICOS)
def cargar_base_precios_completa(gc):
    try:
        url_precios = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
        sh_precios = gc.open_by_url(url_precios)
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
                        if "DOSIS" in str(row[col_precio_tipo]).upper():
                            continue
                            
                    if len(row) > max(col_anio, col_prod):
                        anio_str = str(row[col_anio]).strip()
                        str_prod = str(row[col_prod]).strip().upper()
                        if anio_str.isdigit() and str_prod:
                            col_inicio = max(col_anio, col_prod) + 1
                            vals = [parsear_precio(v) for v in row[col_inicio:] if str(v).strip() != ""]
                            vals = [v for v in vals if v > 0]
                            prom = sum(vals)/len(vals) if vals else 0.0
                            if prom > 0:
                                precios_consolidados.append({'AÑO': int(anio_str), 'PRODUCTO': str_prod, 'PRECIO': prom})
        return pd.DataFrame(precios_consolidados)
    except:
        return pd.DataFrame()

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-presupuesto { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    div[data-testid="stDataFrame"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    .kpi-presupuesto { background-color: #0d1b2a; color: white; padding: 20px; border-radius: 10px; border-left: 6px solid #d4af37; box-shadow: 0 4px 6px rgba(0,0,0,0.2); }
    .kpi-titulo { color: #d4af37; font-weight: bold; font-size: 14px; margin-bottom: 5px; text-transform: uppercase; }
    .kpi-valor { font-size: 32px; font-weight: 900; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-presupuesto'>💰 Módulo 14: Pronóstico Financiero</h1>", unsafe_allow_html=True)
    st.write("Proyección estratégica del flujo de efectivo con rastreo de precios históricos y ajuste de inflación.")

    # --- CONTROLES DE MANDO ---
    st.markdown("### ⚙️ Parámetros del Presupuesto")
    col1, col2, col3, col4, col5 = st.columns([1.5, 1, 1.2, 1, 1.2])
    
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    opciones_mes = ["📊 AÑO COMPLETO (TODOS)"] + list(meses_dict.values())
    
    mes_sel = col1.selectbox("📅 Período:", opciones_mes)
    pista_sel = col2.selectbox("📍 Base:", ["TODAS", "PLUC", "PORI", "PDIV", "TEHO", "LUCI"])
    profundidad_sel = col3.selectbox("🔍 Histórico:", ["Último Año", "Últimos 2 Años", "Últimos 3 Años", "Histórico Completo"])
    crecimiento_sel = col4.number_input("📈 Ha Crecimiento (%)", min_value=-50, max_value=200, value=0, step=5)
    inflacion_sel = col5.number_input("💸 Inflación Anual Estimada (%)", min_value=0.0, max_value=50.0, value=6.0, step=1.0)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 CALCULAR PRESUPUESTO FINANCIERO", type="primary", use_container_width=True):
        with st.spinner("Ejecutando Triage de Precios: Buscando en Año 2026..."):
            try:
                gc = inicializar_cliente_gspread()
                boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                
                df_t1 = pd.DataFrame(boveda.worksheet("TABLA 1").get_all_values()[5:], columns=[str(c).upper().strip() for c in boveda.worksheet("TABLA 1").get_all_values()[4]])
                df_mezclas = pd.DataFrame(boveda.worksheet("DD_Mesclas").get_all_values()[1:], columns=[str(c).upper().strip() for c in boveda.worksheet("DD_Mesclas").get_all_values()[0]])
                
                cfg_data = boveda.worksheet("Configuración").get_all_values()
                df_cfg = pd.DataFrame(cfg_data[1:], columns=[str(c).upper().strip() for c in cfg_data[0]])

                # Descargar base de datos maestra de precios
                df_precios_master = cargar_base_precios_completa(gc)
                dict_precios_backup = extraer_precios_maestros(df_cfg)

                dict_fert = {}
                if len(df_mezclas.columns) > 13:
                    for _, row in df_mezclas.iterrows():
                        f_n = str(row.iloc[12]).strip().upper() 
                        f_s = str(row.iloc[13]).strip().upper() 
                        if f_s and f_n not in ["", "NAN", "NONE", "FERTILIZANTES", "SIGLAS"]:
                            dict_fert[f_s] = f_n

                col_fecha = next((c for c in df_t1.columns if 'FECHA' in c), 'FECHA')
                col_ha = next((c for c in df_t1.columns if 'NETA' in c or 'FUMIG' in c or 'HECT' in c), None)
                col_coctel = next((c for c in df_t1.columns if 'COCTEL' in c or 'CÓCTEL' in c or 'MEZCLA' in c), None)
                col_pista_t1 = next((c for c in df_t1.columns if 'PISTA' in c or 'BASE' in c), None)

                df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
                df_t1 = df_t1.dropna(subset=['FECHA_DT'])
                df_t1['MES'] = df_t1['FECHA_DT'].dt.month
                df_t1['AÑO'] = df_t1['FECHA_DT'].dt.year
                df_t1['HA_CALCULO'] = df_t1[col_ha].apply(a_numero_limpio)
                df_t1['PISTA_OPERATIVA'] = df_t1[col_pista_t1].astype(str).str.upper().str.strip()

                año_actual_operacion = 2026 # Forzado por calendario operativo
                if profundidad_sel == "Último Año": df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 1)]
                elif profundidad_sel == "Últimos 2 Años": df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 2)]
                elif profundidad_sel == "Últimos 3 Años": df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 3)]

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
                    ha_total_por_coctel = df_t1.groupby(['PISTA_OPERATIVA', col_coctel])['HA_CALCULO'].sum().reset_index()
                    factor_crecimiento = 1 + (crecimiento_sel / 100.0)
                    ha_total_por_coctel['HA_PROYECTADA'] = (ha_total_por_coctel['HA_CALCULO'] / total_anios_boveda) * factor_crecimiento

                    for _, row_c in ha_total_por_coctel.iterrows():
                        coctel_completo = str(row_c[col_coctel]).upper().strip()
                        ha_proyectada = row_c['HA_proyectada'] if 'HA_proyectada' in row_c else row_c['HA_PROYECTADA']

                        receta_dict = extraer_receta_completa(coctel_completo, df_mezclas, dict_fert)
                        for prod_quimico, dosis in receta_dict.items():
                            consumo_esperado[prod_quimico] = consumo_esperado.get(prod_quimico, 0) + (dosis * ha_proyectada)

                resultados = []
                gran_total_presupuesto = 0.0

                for producto, volumen in consumo_esperado.items():
                    if volumen > 0:
                        precio_unitario = 0.0
                        anio_origen = ""
                        p_clean = re.sub(r'[^\w]', '', producto.upper().strip())
                        
                        # 💥 1. CRITERIO DE PRIORIDAD ALFA: BUSCAR PRIMERO EN AÑO 2026 (Pestaña Datos) 💥
                        if not df_precios_master.empty:
                            df_2026 = df_precios_master[df_precios_master['AÑO'] == 2026]
                            for _, r_db in df_2026.iterrows():
                                db_clean = re.sub(r'[^\w]', '', str(r_db['PRODUCTO']).upper().strip())
                                if p_clean in db_clean or db_clean in p_clean:
                                    precio_unitario = r_db['PRECIO']
                                    anio_origen = "Base 2026 (Actual)"
                                    break
                        
                        # 💥 2. CRITERIO DE PRIORIDAD BETA: SI NO ESTÁ, PASAR AL HISTÓRICO 💥
                        if precio_unitario == 0.0 and not df_precios_master.empty:
                            df_hist = df_precios_master[df_precios_master['AÑO'] < 2026]
                            matches_hist = []
                            for _, r_db in df_hist.iterrows():
                                db_clean = re.sub(r'[^\w]', '', str(r_db['PRODUCTO']).upper().strip())
                                if p_clean in db_clean or db_clean in p_clean:
                                    matches_hist.append(r_db)
                            
                            if matches_hist:
                                best_match = max(matches_hist, key=lambda x: x['AÑO'])
                                anio_hist = int(best_match['AÑO'])
                                precio_hist = best_match['PRECIO']
                                
                                anios_pasados = max(0, 2026 - anio_hist)
                                precio_unitario = precio_hist * ((1 + (inflacion_sel / 100.0)) ** anios_pasados)
                                anio_origen = f"Rescatado {anio_hist} (+{anios_pasados} años al {inflacion_sel}%)"

                        # 3. ÚLTIMO RECURSO SEGURIDAD: Configuración Local (Solo si es un valor lógico válido)
                        if precio_unitario == 0.0:
                            precio_bk = dict_precios_backup.get(producto, 0.0)
                            if precio_bk < 1000:
                                for p_bk, val_bk in dict_precios_backup.items():
                                    bk_clean = re.sub(r'[^\w]', '', p_bk.upper().strip())
                                    if p_clean in bk_clean or bk_clean in p_clean:
                                        if val_bk >= 1000: precio_unitario = val_bk; break
                            else:
                                precio_unitario = precio_bk
                            if precio_unitario >= 1000:
                                anio_origen = "Configuración Local"

                        if precio_unitario == 0.0:
                            anio_origen = "⚠️ Falta Precio"

                        presupuesto_prod = volumen * precio_unitario
                        gran_total_presupuesto += presupuesto_prod
                        
                        resultados.append({
                            "🧪 INSUMO QUÍMICO": producto,
                            "📦 VOLUMEN ESTIMADO (L/Kg)": volumen,
                            "💵 COSTO UNITARIO ($)": precio_unitario,
                            "🔎 ORIGEN PRECIO": anio_origen,
                            "💰 PRESUPUESTO ASIGNADO ($)": presupuesto_prod
                        })

                df_presupuesto = pd.DataFrame(resultados)

                st.markdown("---")
                
                if df_presupuesto.empty:
                    st.warning("⚠️ No hay suficientes datos históricos para proyectar este escenario.")
                else:
                    st.markdown(f"""
                    <div class='kpi-presupuesto'>
                        <div class='kpi-titulo'>FLUJO DE EFECTIVO PROYECTADO ({mes_sel})</div>
                        <p class='kpi-valor'>$ {fmt_latino(gran_total_presupuesto, 0)} <span style='font-size: 16px; font-weight: normal;'>COP</span></p>
                        <p style='margin: 0; margin-top: 10px; color: #a0aec0; font-size: 12px;'>Calculado sobre {fmt_latino(ha_total_detectada/total_anios_boveda * (1 + crecimiento_sel/100))} Hectáreas proyectadas. Factor Crecimiento: {crecimiento_sel}% | Inflación de ajuste: {inflacion_sel}%</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br>### 📋 Desglose Financiero por Insumo", unsafe_allow_html=True)
                    
                    # Ordenamiento Estricto Alfabético
                    df_presupuesto = df_presupuesto.sort_values(by="🧪 INSUMO QUÍMICO", ascending=True)
                    
                    df_vista = df_presupuesto.copy()
                    df_vista['📦 VOLUMEN ESTIMADO (L/Kg)'] = df_vista['📦 VOLUMEN ESTIMADO (L/Kg)'].apply(lambda x: fmt_latino(x, 1))
                    
                    def formatear_precio(x):
                        return f"$ {fmt_latino(x, 0)}" if x > 0 else "⚠️ Falta Dato"

                    df_vista['💵 COSTO UNITARIO ($)'] = df_vista['💵 COSTO UNITARIO ($)'].apply(formatear_precio)
                    df_vista['💰 PRESUPUESTO ASIGNADO ($)'] = df_vista['💰 PRESUPUESTO ASIGNADO ($)'].apply(lambda x: f"$ {fmt_latino(x, 0)}" if x > 0 else "$ 0")

                    st.dataframe(
                        df_vista, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "🧪 INSUMO QUÍMICO": st.column_config.TextColumn("INSUMO QUÍMICO", width="medium"),
                            "📦 VOLUMEN ESTIMADO (L/Kg)": st.column_config.TextColumn("VOLUMEN REQUERIDO", width="small"),
                            "💵 COSTO UNITARIO ($)": st.column_config.TextColumn("COSTO UNITARIO", width="small"),
                            "🔎 ORIGEN PRECIO": st.column_config.TextColumn("ORIGEN DEL DATO", width="medium"),
                            "💰 PRESUPUESTO ASIGNADO ($)": st.column_config.TextColumn("PRESUPUESTO TOTAL", width="medium")
                        }
                    )
            except Exception as e:
                st.error(f"🚨 Falla en los cálculos financieros: {e}")
