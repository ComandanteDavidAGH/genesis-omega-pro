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

# 🧠 CEREBRO QUÍMICO CALIBRADO (SIN INYECCIÓN FANTASMA DE ADITIVOS)
def extraer_receta_completa(coctel_sel, df_mezclas, dict_fertilizantes_dinamico):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if len(partes) > 0 else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    
    # 1. Base del Cóctel (Lectura Estricta)
    if not df_mezclas.empty:
        col_0_limpia = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0_limpia == base_coctel]
        for _, r in rb.iterrows():
            p = str(r.iloc[1]).strip().upper()
            d = a_numero_limpio(r.iloc[2])
            if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    # 2. Inyección Dinámica de Fertilizantes (Solo los que pida la Sigla)
    for aditivo in aditivos:
        if aditivo in dict_fertilizantes_dinamico:
            nombre_fert = dict_fertilizantes_dinamico[aditivo]
            dosis_fert = obtener_dosis_fertilizante(df_mezclas, nombre_fert)
            
            # Solo inyecta el producto SI encuentra una dosis válida en la matriz
            if dosis_fert is not None:
                dict_prods[nombre_fert] = dict_prods.get(nombre_fert, 0.0) + dosis_fert
            elif aditivo == "NM": dict_prods["NATURAMIN WSP"] = 0.2
            elif aditivo == "ZN": dict_prods["ZINTRAC X LITRO SV"] = 0.5
            elif aditivo == "BT": dict_prods["BANATREL SC"] = 0.5
    
    # 3. Aditivos Universales Estrictos
    if "SV" in coctel_u or "ACONDICIONADOR" in coctel_u:
        dict_prods["ACONDICIONADOR SV"] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
        dict_prods["ADHERENTE SV"] = 0.13
        
    if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel: 
        dict_prods["IMBIOSIL O"] = 1.5

    return dict_prods

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-oraculo { color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }
    .alerta-roja { background-color: #ffe6e6; color: #cc0000; padding: 15px; border-left: 8px solid #cc0000; border-radius: 5px; font-weight: bold; margin-bottom: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    .alerta-amarilla { background-color: #fff3cd; color: #856404; padding: 15px; border-left: 8px solid #ffc107; border-radius: 5px; font-weight: bold; margin-bottom: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    .alerta-verde { background-color: #d4edda; color: #155724; padding: 15px; border-left: 8px solid #28a745; border-radius: 5px; font-weight: bold; margin-bottom: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    div[data-testid="stDataFrame"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-oraculo'>🔮 El Oráculo: Predicción Cíclica de Rupturas</h1>", unsafe_allow_html=True)
    st.write("Análisis estacional del comportamiento epidemiológico cruzado con inventarios de SAP.")

    st.markdown("### 📥 1. Radar de Existencias Actuales (SAP)")
    archivo_sap = st.file_uploader("Cargue la Sábana SAP actualizada (.xlsx o .csv)", type=['xlsx', 'csv'], key="sap_oraculo")
    
    st.markdown("### 📅 2. Parámetros de Predicción")
    col_mes, col_pista, col_profundidad = st.columns([1.2, 1.2, 1.5])
    
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    mes_actual = datetime.now().month
    mes_proyeccion = col_mes.selectbox("Mes a Proyectar:", list(meses_dict.keys()), index=mes_actual-1, format_func=lambda x: meses_dict[x])
    
    # 💥 LISTA PURGADA (Se eliminaron Z-1 y Z-2) 💥
    lista_pistas = ["TODAS", "PLUC", "PORI", "PDIV", "TEHO", "LUCI"]
    pista_objetivo = col_pista.selectbox("📍 Base Operativa (SAP):", lista_pistas)

    opciones_profundidad = ["Último Año (Tendencia Reciente)", "Últimos 2 Años", "Últimos 3 Años", "Histórico Completo"]
    profundidad_sel = col_profundidad.selectbox("🔍 Profundidad del Histórico:", opciones_profundidad)

    st.markdown("<br>", unsafe_allow_html=True)

    if not archivo_sap:
        st.info("💡 Despliegue el archivo SAP para que el sistema evalúe el blindaje de las pistas.")
        return

    if st.button("🚀 EJECUTAR PREDICCIÓN ESTACIONAL", type="primary", use_container_width=True):
        with st.spinner(f"Sincronizando lenguajes y analizando comportamiento agronómico..."):
            try:
                # --- A. LECTURA DE SAP ---
                if archivo_sap.name.lower().endswith('.xlsx') or archivo_sap.name.lower().endswith('.xls'):
                    df_sap = pd.read_excel(archivo_sap)
                else:
                    try: df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='utf-8')
                    except:
                        archivo_sap.seek(0)
                        df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='latin1')

                def purificar_columna(col_name):
                    return str(col_name).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').strip()
                
                cols_limpias = [purificar_columna(c) for c in df_sap.columns]
                
                idx_cod = next((i for i, c in enumerate(cols_limpias) if 'MATERIAL' in c or 'COD' in c or 'ITEM' in c), None)
                idx_prod = next((i for i, c in enumerate(cols_limpias) if ('TEXTO' in c or 'DESC' in c or 'PRODUCTO' in c or 'DENOMINACION' in c) and i != idx_cod), None)
                idx_pista = next((i for i, c in enumerate(cols_limpias) if 'ALMACEN' in c or 'PISTA' in c or 'LGORT' in c), None)
                idx_saldo = next((i for i, c in enumerate(cols_limpias) if 'LIBRE' in c or 'SALDO' in c or 'UTILIZACION' in c or 'LABST' in c), None)

                if idx_prod is None or idx_pista is None or idx_saldo is None:
                    st.error(f"❌ Error de Radar: No se pudieron mapear las columnas críticas en SAP. Columnas detectadas: {list(df_sap.columns)}")
                    return
                
                c_prod = df_sap.columns[idx_prod]
                c_pista = df_sap.columns[idx_pista]
                c_saldo = df_sap.columns[idx_saldo]
                c_cod = df_sap.columns[idx_cod] if idx_cod is not None else None

                if c_cod is not None and c_prod is not None:
                    df_sap['PRODUCTO_RADAR'] = df_sap[c_cod].astype(str).str.split('.').str[0].str.strip() + " | " + df_sap[c_prod].astype(str).str.upper().str.strip()
                else:
                    df_sap['PRODUCTO_RADAR'] = df_sap[c_prod].astype(str).str.upper().str.strip()

                df_sap['SALDO_FISICO'] = df_sap[c_saldo].apply(a_numero_limpio)
                df_sap['PISTA_SAP'] = df_sap[c_pista].astype(str).str.upper().str.strip()
                
                df_sap_agrupado = df_sap.groupby(['PISTA_SAP', 'PRODUCTO_RADAR'])['SALDO_FISICO'].sum().reset_index()
                df_sap_agrupado = df_sap_agrupado[df_sap_agrupado['SALDO_FISICO'] > 0]

                if pista_objetivo != "TODAS":
                    df_sap_agrupado = df_sap_agrupado[df_sap_agrupado['PISTA_SAP'].str.contains(pista_objetivo, na=False)]

                # --- B. LECTURA DE BÓVEDA ---
                gc = inicializar_cliente_gspread()
                boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                
                df_t1 = pd.DataFrame(boveda.worksheet("TABLA 1").get_all_values()[5:], columns=[str(c).upper().strip() for c in boveda.worksheet("TABLA 1").get_all_values()[4]])
                df_mezclas = pd.DataFrame(boveda.worksheet("DD_Mesclas").get_all_values()[1:], columns=[str(c).upper().strip() for c in boveda.worksheet("DD_Mesclas").get_all_values()[0]])

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

                año_actual_operacion = datetime.now().year
                if profundidad_sel == "Último Año (Tendencia Reciente)":
                    df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 1)]
                elif profundidad_sel == "Últimos 2 Años":
                    df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 2)]
                elif profundidad_sel == "Últimos 3 Años":
                    df_t1 = df_t1[df_t1['AÑO'] >= (año_actual_operacion - 3)]
                
                max_date = df_t1['FECHA_DT'].max()
                fecha_limite = max_date - timedelta(days=90)
                df_reciente = df_t1[df_t1['FECHA_DT'] >= fecha_limite]
                
                ha_mensual_actual_pista = {}
                if not df_reciente.empty:
                    ha_mensual_actual_pista = (df_reciente.groupby('PISTA_OPERATIVA')['HA_CALCULO'].sum() / 3.0).to_dict()

                df_hist_mes = df_t1[df_t1['MES'] == mes_proyeccion].copy()
                consumo_esperado_pista = {} 
                ha_total_detectada = 0.0

                if not df_hist_mes.empty:
                    ha_total_detectada = df_hist_mes['HA_CALCULO'].sum()
                    ha_hist_total_pista = df_hist_mes.groupby('PISTA_OPERATIVA')['HA_CALCULO'].sum().to_dict()

                    volumen_hist_total = {}
                    for _, row_c in df_hist_mes.iterrows():
                        pista_op = row_c['PISTA_OPERATIVA']
                        coctel_completo = str(row_c[col_coctel]).upper().strip()
                        ha_aplicadas = row_c['HA_CALCULO']
                        
                        if pista_op not in volumen_hist_total:
                            volumen_hist_total[pista_op] = {}

                        receta_dict = extraer_receta_completa(coctel_completo, df_mezclas, dict_fert)
                        for prod_quimico, dosis in receta_dict.items():
                            volumen_hist_total[pista_op][prod_quimico] = volumen_hist_total[pista_op].get(prod_quimico, 0) + (dosis * ha_aplicadas)

                    # Fusión Híbrida
                    for pista_op, prods in volumen_hist_total.items():
                        ha_historicas_mes = ha_hist_total_pista.get(pista_op, 0)
                        ha_actuales_mes = ha_mensual_actual_pista.get(pista_op, ha_historicas_mes / df_hist_mes['AÑO'].nunique())
                        
                        if pista_op not in consumo_esperado_pista:
                            consumo_esperado_pista[pista_op] = {}
                            
                        for prod, vol_hist in prods.items():
                            if ha_historicas_mes > 0:
                                dosis_promedio_blended = vol_hist / ha_historicas_mes
                                consumo_esperado_pista[pista_op][prod] = dosis_promedio_blended * ha_actuales_mes
                            else:
                                consumo_esperado_pista[pista_op][prod] = 0.0

                traductor_pistas = {"PLUC": "FUMIGARAY", "PORI": "AEROPENOR", "LUCI": "GENESYS", "TEHO": "AVIL", "PDIV": "ASA"}

                resultados = []
                for _, row_s in df_sap_agrupado.iterrows():
                    pista_sap = row_s['PISTA_SAP']
                    producto_sap_completo = str(row_s['PRODUCTO_RADAR']).upper().strip()
                    saldo = row_s['SALDO_FISICO']

                    consumo_mes_proyectado = 0.0
                    pista_t1_esperada = traductor_pistas.get(pista_sap, pista_sap)
                    pista_clave = next((k for k in consumo_esperado_pista.keys() if pista_t1_esperada in k or k in pista_t1_esperada), None)
                    
                    if pista_clave:
                        for p_receta, vol_mes in consumo_esperado_pista[pista_clave].items():
                            p_receta_clean = p_receta.replace(" ", "")
                            prod_sap_clean = producto_sap_completo.replace(" ", "")
                            if p_receta_clean in prod_sap_clean or prod_sap_clean in p_receta_clean:
                                consumo_mes_proyectado += vol_mes

                    consumo_diario = consumo_mes_proyectado / 30 if consumo_mes_proyectado > 0 else 0
                    
                    if consumo_diario > 0:
                        dias_autonomia = saldo / consumo_diario
                        if dias_autonomia <= 7: estado = "🚨 CRÍTICO (< 7 Días)"
                        elif dias_autonomia <= 21: estado = "⚠️ ALERTA (8-21 Días)"
                        else: estado = "✅ ÓPTIMO (> 21 Días)"
                    else:
                        dias_autonomia = 9999
                        estado = "✅ ÓPTIMO (Sin Consumo Histórico)"

                    resultados.append({
                        "📍 PISTA": pista_sap,
                        "🧪 CÓDIGO | PRODUCTO": producto_sap_completo,
                        "📦 SALDO (SAP)": saldo,
                        "📈 PROYECCIÓN MES (L/Kg)": round(consumo_mes_proyectado, 1),
                        "⏳ AUTONOMÍA": round(dias_autonomia, 0),
                        "ESTADO": estado
                    })

                df_oraculo = pd.DataFrame(resultados)

                st.markdown("---")
                
                if ha_total_detectada > 0:
                    st.success(f"✅ Motor Híbrido Activado: El sistema analizó el patrón químico histórico de {meses_dict[mes_proyeccion]} y lo ajustó automáticamente al crecimiento en hectáreas de los últimos 90 días.")
                else:
                    st.warning(f"⚠️ El radar no encontró hectáreas operadas en el mes de {meses_dict[mes_proyeccion]} dentro de la base histórica seleccionada.")

                st.markdown(f"### 🎯 Tablero Táctico: Proyección para {meses_dict[mes_proyeccion]}")
                
                if df_oraculo.empty:
                    st.info("No se hallaron productos en SAP para la pista seleccionada.")
                else:
                    # 💥 ORDENAMIENTO ALFABÉTICO PERFECTO 💥
                    def get_sort_weight(estado_str):
                        if "CRÍTICO" in estado_str: return 1
                        if "ALERTA" in estado_str: return 2
                        return 3

                    df_oraculo['SORT_WEIGHT'] = df_oraculo['ESTADO'].apply(get_sort_weight)
                    df_oraculo['SOLO_NOMBRE'] = df_oraculo['🧪 CÓDIGO | PRODUCTO'].apply(lambda x: x.split('|')[1].strip() if '|' in x else x)
                    
                    df_oraculo = df_oraculo.sort_values(by=["📍 PISTA", "SORT_WEIGHT", "SOLO_NOMBRE"], ascending=[True, True, True])
                    df_oraculo = df_oraculo.drop(columns=['SORT_WEIGHT', 'SOLO_NOMBRE'])
                    
                    criticos = len(df_oraculo[df_oraculo['ESTADO'] == "🚨 CRÍTICO (< 7 Días)"])
                    alertas = len(df_oraculo[df_oraculo['ESTADO'] == "⚠️ ALERTA (8-21 Días)"])
                    optimos = len(df_oraculo) - (criticos + alertas)
                    
                    c_k1, c_k2, c_k3 = st.columns(3)
                    
                    c_k1.markdown(f"""
                    <div style="background-color: #ffe6e6; border-left: 5px solid #cc0000; padding: 10px; border-radius: 5px;">
                        <span style="color: #cc0000; font-weight: bold;">🚨 RUPTURA INMINENTE</span><br/>
                        <span style="font-size: 18px; color: #0d1b2a; font-weight: bold;">{criticos} Insumos</span>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    c_k2.markdown(f"""
                    <div style="background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 10px; border-radius: 5px;">
                        <span style="color: #856404; font-weight: bold;">⚠️ ALERTA LOGÍSTICA</span><br/>
                        <span style="font-size: 18px; color: #0d1b2a; font-weight: bold;">{alertas} Insumos</span>
                    </div>
                    """, unsafe_allow_html=True)

                    c_k3.markdown(f"""
                    <div style="background-color: #d4edda; border-left: 5px solid #28a745; padding: 10px; border-radius: 5px;">
                        <span style="color: #155724; font-weight: bold;">✅ INVENTARIO SANO</span><br/>
                        <span style="font-size: 18px; color: #0d1b2a; font-weight: bold;">{optimos} Insumos</span>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br/>", unsafe_allow_html=True)
                    
                    def pintar_oraculo(row):
                        if "CRÍTICO" in row['ESTADO']: return ['background-color: #ffe6e6; color: #cc0000; font-weight:bold;'] * len(row)
                        if "ALERTA" in row['ESTADO']: return ['background-color: #fff3cd; color: #856404; font-weight:bold;'] * len(row)
                        return [''] * len(row)

                    df_vista = df_oraculo.copy()
                    df_vista['📦 SALDO (SAP)'] = df_vista['📦 SALDO (SAP)'].apply(lambda x: fmt_latino(x, 1))
                    df_vista['📈 PROYECCIÓN MES (L/Kg)'] = df_vista['📈 PROYECCIÓN MES (L/Kg)'].apply(lambda x: fmt_latino(x, 1))
                    df_vista['⏳ AUTONOMÍA'] = df_vista['⏳ AUTONOMÍA'].apply(lambda x: "∞ (Sin Consumo Histórico)" if x >= 9999 else f"{x:,.0f} Días")

                    # 💥 REDERIZADO NATIVO PARA ALINEACIÓN PERFECTA DE COLUMNAS 💥
                    st.dataframe(
                        df_vista.style.apply(pintar_oraculo, axis=1), 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "📍 PISTA": st.column_config.TextColumn("PISTA", width="small"),
                            "🧪 CÓDIGO | PRODUCTO": st.column_config.TextColumn("PRODUCTO", width="large"),
                            "📦 SALDO (SAP)": st.column_config.TextColumn("SALDO (SAP)", width="medium"),
                            "📈 PROYECCIÓN MES (L/Kg)": st.column_config.TextColumn("PROYECCIÓN MES", width="medium"),
                            "⏳ AUTONOMÍA": st.column_config.TextColumn("AUTONOMÍA", width="medium"),
                            "ESTADO": st.column_config.TextColumn("ESTADO", width="medium")
                        }
                    )
                    
            except Exception as e:
                st.error(f"🚨 Falla en los cálculos predictivos o estructura de datos: {e}")
