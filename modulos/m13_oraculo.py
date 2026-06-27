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

# 🧠 CEREBRO QUÍMICO CALIBRADO
def extraer_receta_completa(coctel_sel, df_mezclas, dict_fertilizantes_dinamico):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if len(partes) > 0 else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    
    # 1. Base del Cóctel
    if not df_mezclas.empty:
        col_0_limpia = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0_limpia == base_coctel]
        for _, r in rb.iterrows():
            p = str(r.iloc[1]).strip().upper()
            d = a_numero_limpio(r.iloc[2])
            if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    # 2. Inyección Dinámica de Fertilizantes
    for aditivo in aditivos:
        if aditivo in dict_fertilizantes_dinamico:
            nombre_fert = dict_fertilizantes_dinamico[aditivo]
            dosis_fert = obtener_dosis_fertilizante(df_mezclas, nombre_fert)
            
            # Salvavidas calibrado a la realidad agronómica
            if dosis_fert is None:
                if "NATURAMIN" in nombre_fert: dosis_fert = 0.2
                elif "ZINTRAC" in nombre_fert: dosis_fert = 0.5
                elif "BANATREL" in nombre_fert: dosis_fert = 0.5
                else: dosis_fert = 0.5
                
            dict_prods[nombre_fert] = dict_prods.get(nombre_fert, 0.0) + dosis_fert
    
    # 3. Aditivos Universales
    if not any("ADHERENTE" in k for k in dict_prods.keys()): dict_prods["ADHERENTE SV"] = 0.13
    if not any("ACONDICIONADOR" in k for k in dict_prods.keys()): 
        dict_prods["ACONDICIONADOR SV"] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
    if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel: 
        dict_prods["IMBIOSIL O"] = 1.5

    return dict_prods

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-oraculo { color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-oraculo'>🔮 El Oráculo: Predicción Cíclica de Rupturas</h1>", unsafe_allow_html=True)
    st.write("Análisis estacional del comportamiento epidemiológico cruzado con inventarios de SAP.")

    st.markdown("### 📥 1. Radar de Existencias Actuales (SAP)")
    archivo_sap = st.file_uploader("Cargue la Sábana SAP actualizada (.xlsx o .csv)", type=['xlsx', 'csv'], key="sap_oraculo")
    
    st.markdown("### 📅 2. Parámetros de Predicción")
    col_mes, col_pista, col_vacia = st.columns([1.5, 1.5, 1])
    
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    mes_actual = datetime.now().month
    mes_proyeccion = col_mes.selectbox("Mes a Proyectar (Ciclo Histórico):", list(meses_dict.keys()), index=mes_actual-1, format_func=lambda x: meses_dict[x])
    
    lista_pistas = ["TODAS", "PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2"]
    pista_objetivo = col_pista.selectbox("📍 Filtrar por Pista Operativa (SAP):", lista_pistas)

    st.markdown("<br>", unsafe_allow_html=True)

    if not archivo_sap:
        st.info("💡 Despliegue el archivo SAP para que el sistema evalúe el blindaje de las pistas.")
        return

    if st.button("🚀 EJECUTAR PREDICCIÓN ESTACIONAL", type="primary", use_container_width=True):
        with st.spinner(f"Sincronizando lenguajes (SAP vs Operaciones) y analizando años anteriores..."):
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

                # --- B. LECTURA DE HISTÓRICO Y RECETAS (BÓVEDA) ---
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

                # RECUPERACIÓN DE FECHAS PESADA
                df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
                df_t1 = df_t1.dropna(subset=['FECHA_DT'])
                df_t1['MES'] = df_t1['FECHA_DT'].dt.month
                df_t1['AÑO'] = df_t1['FECHA_DT'].dt.year
                
                total_anios_boveda = df_t1['AÑO'].nunique()
                if total_anios_boveda == 0: total_anios_boveda = 1

                df_hist_mes = df_t1[df_t1['MES'] == mes_proyeccion].copy()
                
                consumo_esperado_pista = {} 
                ha_total_detectada = 0.0

                if not df_hist_mes.empty:
                    df_hist_mes['HA_CALCULO'] = df_hist_mes[col_ha].apply(a_numero_limpio)
                    df_hist_mes['PISTA_OPERATIVA'] = df_hist_mes[col_pista_t1].astype(str).str.upper().str.strip()
                    ha_total_detectada = df_hist_mes['HA_CALCULO'].sum()

                    ha_total_por_coctel = df_hist_mes.groupby(['PISTA_OPERATIVA', col_coctel])['HA_CALCULO'].sum().reset_index()
                    ha_total_por_coctel['HA_PROMEDIO'] = ha_total_por_coctel['HA_CALCULO'] / total_anios_boveda

                    for _, row_c in ha_total_por_coctel.iterrows():
                        pista_op = row_c['PISTA_OPERATIVA']
                        coctel_completo = str(row_c[col_coctel]).upper().strip()
                        ha_promedio_aplicadas = row_c['HA_PROMEDIO']
                        
                        if pista_op not in consumo_esperado_pista:
                            consumo_esperado_pista[pista_op] = {}

                        receta_dict = extraer_receta_completa(coctel_completo, df_mezclas, dict_fert)
                        for prod_quimico, dosis in receta_dict.items():
                            consumo_esperado_pista[pista_op][prod_quimico] = consumo_esperado_pista[pista_op].get(prod_quimico, 0) + (dosis * ha_promedio_aplicadas)

                traductor_pistas = {
                    "PLUC": "FUMIGARAY", "PORI": "AEROPENOR", "LUCI": "GENESYS", "TEHO": "AVIL", "PDIV": "ASA"
                }

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
                    st.success(f"✅ Memoria Histórica Recuperada: El radar evaluó {total_anios_boveda} años de historia y promedió un volumen de **{fmt_latino(ha_total_detectada / total_anios_boveda)} Ha/Año** para el mes de {meses_dict[mes_proyeccion]}.")
                else:
                    st.warning(f"⚠️ El radar no encontró hectáreas operadas en el mes de {meses_dict[mes_proyeccion]} en su base de datos histórica.")

                st.markdown(f"### 🎯 Tablero Táctico: Proyección para {meses_dict[mes_proyeccion]}")
                
                if df_oraculo.empty:
                    st.info("No se hallaron productos en SAP para la pista seleccionada.")
                else:
                    # 💥 ALGORITMO TRIAGE CORREGIDO (Solo 3 Niveles Reales) 💥
                    def get_sort_weight(estado_str):
                        if "CRÍTICO" in estado_str: return 1
                        if "ALERTA" in estado_str: return 2
                        return 3 # Agrupa a todos los "Verdes" (Tanto > 21 Días como los Sin Consumo) para que se ordenen juntos de la A a la Z

                    df_oraculo['SORT_WEIGHT'] = df_oraculo['ESTADO'].apply(get_sort_weight)
                    
                    # Ordena primero por pista, luego por prioridad (Rojo, Amarillo, Verde), y luego ESTRICTAMENTE ALFABÉTICO por código/producto
                    df_oraculo = df_oraculo.sort_values(by=["📍 PISTA", "SORT_WEIGHT", "🧪 CÓDIGO | PRODUCTO"], ascending=[True, True, True])
                    df_oraculo = df_oraculo.drop(columns=['SORT_WEIGHT'])
                    
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

                    st.dataframe(df_vista.style.apply(pintar_oraculo, axis=1), use_container_width=True, hide_index=True)
                    
            except Exception as e:
                st.error(f"🚨 Falla en los cálculos predictivos o estructura de datos: {e}")
