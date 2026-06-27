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

# 💥 TRADUCTOR MÉTRICO LATINO
def fmt_latino(val, decimales=1):
    try: return f"{float(val):,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(val)

# --- 🧠 CEREBRO DEL ORÁCULO CÍCLICO ---
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
    st.write("Análisis estacional del comportamiento de la Sigatoka cruzado con inventarios actuales de SAP segregados por pista operativa.")

    # 1. CARGA DEL INVENTARIO ACTUAL
    st.markdown("### 📥 1. Radar de Existencias Actuales (SAP)")
    archivo_sap = st.file_uploader("Cargue la Sábana SAP actualizada (.xlsx o .csv)", type=['xlsx', 'csv'], key="sap_oraculo")
    
    # 2. SELECTOR EPIDEMIOLÓGICO (MES OBJETIVO Y PISTA)
    st.markdown("### 📅 2. Parámetros de Predicción")
    col_mes, col_pista, col_vacia = st.columns([1.5, 1.5, 1])
    
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    mes_actual = datetime.now().month
    
    mes_proyeccion = col_mes.selectbox("Mes a Proyectar (Ciclo Histórico):", list(meses_dict.keys()), index=mes_actual-1, format_func=lambda x: meses_dict[x])
    
    # Lista estandarizada de pistas conocidas
    lista_pistas = ["TODAS", "PLUC", "PORI", "PDIV", "TEHO", "LUCI", "Z-1", "Z-2"]
    pista_objetivo = col_pista.selectbox("📍 Filtrar por Pista Operativa:", lista_pistas)

    st.markdown("<br>", unsafe_allow_html=True)

    if not archivo_sap:
        st.info("💡 Despliegue el archivo SAP para que el sistema evalúe el blindaje de las pistas.")
        return

    if st.button("🚀 EJECUTAR PREDICCIÓN ESTACIONAL", type="primary", use_container_width=True):
        with st.spinner(f"Viajando en el tiempo para analizar los ciclos de {meses_dict[mes_proyeccion]} en años anteriores..."):
            try:
                # --- A. LECTURA DE SAP ---
                if archivo_sap.name.lower().endswith('.xlsx') or archivo_sap.name.lower().endswith('.xls'):
                    df_sap = pd.read_excel(archivo_sap)
                else:
                    try: df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='utf-8')
                    except:
                        archivo_sap.seek(0)
                        df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='latin1')

                # 💥 ESCÁNER ROBUSTO DUAL (Código + Nombre)
                def purificar_columna(col_name):
                    return str(col_name).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').strip()
                
                cols_limpias = [purificar_columna(c) for c in df_sap.columns]
                
                c_cod, c_desc, c_pista, c_saldo = None, None, None, None
                for i, c in enumerate(cols_limpias):
                    if ('ALMACEN' in c or 'PISTA' in c or 'LGORT' in c) and not c_pista: c_pista = df_sap.columns[i]
                    elif ('LIBRE' in c or 'SALDO' in c or 'UTILIZACION' in c or 'LABST' in c) and not c_saldo: c_saldo = df_sap.columns[i]
                    elif ('DESC' in c or 'TEXTO' in c or 'PRODUCTO' in c) and not c_desc: c_desc = df_sap.columns[i]
                    elif ('MATERIAL' in c or 'COD' in c or 'ITEM' in c) and not c_cod: c_cod = df_sap.columns[i]

                if not c_saldo or not c_pista or (not c_desc and not c_cod):
                    st.error(f"❌ Error de Radar: No se pudieron mapear las columnas críticas en SAP. (Encontradas: {list(df_sap.columns)})")
                    return
                
                # FUSIÓN ATÓMICA: "CÓDIGO | NOMBRE"
                if c_cod and c_desc and c_cod != c_desc:
                    df_sap['PRODUCTO_RADAR'] = df_sap[c_cod].astype(str).str.split('.').str[0].str.strip() + " | " + df_sap[c_desc].astype(str).str.upper().str.strip()
                elif c_desc:
                    df_sap['PRODUCTO_RADAR'] = df_sap[c_desc].astype(str).str.upper().str.strip()
                else:
                    df_sap['PRODUCTO_RADAR'] = df_sap[c_cod].astype(str).str.split('.').str[0].str.strip()

                df_sap['SALDO_FISICO'] = df_sap[c_saldo].apply(a_numero_limpio)
                df_sap['PISTA_SAP'] = df_sap[c_pista].astype(str).str.upper().str.strip()
                
                df_sap_agrupado = df_sap.groupby(['PISTA_SAP', 'PRODUCTO_RADAR'])['SALDO_FISICO'].sum().reset_index()
                df_sap_agrupado = df_sap_agrupado[df_sap_agrupado['SALDO_FISICO'] > 0]

                # Aplicar Filtro de Pista si no es "TODAS"
                if pista_objetivo != "TODAS":
                    df_sap_agrupado = df_sap_agrupado[df_sap_agrupado['PISTA_SAP'].str.contains(pista_objetivo, na=False)]

                if df_sap_agrupado.empty:
                    st.warning(f"⚠️ No hay inventario disponible en SAP para la pista {pista_objetivo}.")
                    return

                # --- B. LECTURA DE HISTÓRICO (TABLA 1) Y RECETAS ---
                gc = inicializar_cliente_gspread()
                boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                
                t1_data = boveda.worksheet("TABLA 1").get_all_values()
                df_t1 = pd.DataFrame(t1_data[5:], columns=[str(c).upper().strip() for c in t1_data[4]])
                
                mezclas_data = boveda.worksheet("DD_Mesclas").get_all_values()
                df_mezclas = pd.DataFrame(mezclas_data[1:], columns=[str(c).upper().strip() for c in mezclas_data[0]])

                col_fecha = next((c for c in df_t1.columns if 'FECHA' in c), 'FECHA')
                col_ha = next((c for c in df_t1.columns if 'NETA' in c or 'FUMIG' in c or 'HECT' in c), None)
                col_coctel = next((c for c in df_t1.columns if 'COCTEL' in c or 'CÓCTEL' in c or 'MEZCLA' in c), None)
                col_pista_t1 = next((c for c in df_t1.columns if 'PISTA' in c or 'BASE' in c), None)

                if not col_ha or not col_coctel or not col_pista_t1:
                    st.error("🚨 Faltan columnas críticas en TABLA 1 (Hectáreas, Cóctel o Pista).")
                    return

                # Filtrar TABLA 1 solo por el MES seleccionado (Todos los años)
                df_t1['FECHA_DT'] = pd.to_datetime(df_t1[col_fecha], dayfirst=True, errors='coerce')
                df_t1 = df_t1.dropna(subset=['FECHA_DT'])
                df_t1['MES'] = df_t1['FECHA_DT'].dt.month
                df_t1['AÑO'] = df_t1['FECHA_DT'].dt.year
                
                df_hist_mes = df_t1[df_t1['MES'] == mes_proyeccion].copy()
                
                # --- C. CÁLCULO DE PROMEDIO HISTÓRICO ANUAL POR PISTA Y CÓCTEL ---
                consumo_esperado_pista = {} # { PISTA: { INSUMO: VOLUMEN } }
                
                if not df_hist_mes.empty:
                    df_hist_mes['HA_CALCULO'] = df_hist_mes[col_ha].apply(a_numero_limpio)
                    df_hist_mes['PISTA_OPERATIVA'] = df_hist_mes[col_pista_t1].astype(str).str.upper().str.strip()

                    ha_por_anio = df_hist_mes.groupby(['AÑO', 'PISTA_OPERATIVA', col_coctel])['HA_CALCULO'].sum().reset_index()
                    ha_promedio_hist = ha_por_anio.groupby(['PISTA_OPERATIVA', col_coctel])['HA_CALCULO'].mean().reset_index()

                    # --- D. EXPLOSIÓN QUÍMICA ---
                    for _, row_c in ha_promedio_hist.iterrows():
                        pista_op = row_c['PISTA_OPERATIVA']
                        coctel_nombre = str(row_c[col_coctel]).upper().strip().split(" ")[0]
                        ha_aplicadas = row_c['HA_CALCULO']
                        
                        if pista_op not in consumo_esperado_pista:
                            consumo_esperado_pista[pista_op] = {}

                        receta = df_mezclas[df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip() == coctel_nombre]
                        for _, r_mez in receta.iterrows():
                            prod_quimico = str(r_mez.iloc[1]).upper().strip()
                            dosis = a_numero_limpio(r_mez.iloc[2])
                            if prod_quimico not in ["NAN", "", "NONE"] and dosis > 0:
                                consumo_esperado_pista[pista_op][prod_quimico] = consumo_esperado_pista[pista_op].get(prod_quimico, 0) + (dosis * ha_aplicadas)

                # --- E. CRUCE ALGORÍTMICO PISTA vs PISTA ---
                resultados = []
                for _, row_s in df_sap_agrupado.iterrows():
                    pista_sap = row_s['PISTA_SAP']
                    producto_sap = str(row_s['PRODUCTO_RADAR']).upper().strip() # <--- AHORA USA LA FUSIÓN
                    saldo = row_s['SALDO_FISICO']

                    consumo_mes_proyectado = 0.0
                    pista_clave = next((k for k in consumo_esperado_pista.keys() if pista_sap in k or k in pista_sap), None)
                    
                    if pista_clave:
                        for p_receta, vol_mes in consumo_esperado_pista[pista_clave].items():
                            # El producto_sap ahora tiene "101658 | MANCOZEB", así que p_receta ("MANCOZEB") hará match
                            if p_receta in producto_sap or producto_sap in p_receta:
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
                        "🧪 CÓDIGO | PRODUCTO": producto_sap,
                        "📦 SALDO (SAP)": saldo,
                        "📈 PROYECCIÓN MES (L/Kg)": round(consumo_mes_proyectado, 1),
                        "⏳ AUTONOMÍA": round(dias_autonomia, 0),
                        "ESTADO": estado
                    })

                df_oraculo = pd.DataFrame(resultados)

                st.markdown("---")
                st.markdown(f"### 🎯 Tablero Táctico: Proyección para {meses_dict[mes_proyeccion]}")
                
                df_oraculo = df_oraculo.sort_values(by=["ESTADO", "📍 PISTA", "⏳ AUTONOMÍA"])
                
                criticos = len(df_oraculo[df_oraculo['ESTADO'] == "🚨 CRÍTICO (< 7 Días)"])
                alertas = len(df_oraculo[df_oraculo['ESTADO'] == "⚠️ ALERTA (8-21 Días)"])
                optimos = len(df_oraculo) - (criticos + alertas)
                
                c_k1, c_k2, c_k3 = st.columns(3)
                c_k1.markdown(f"<div class='alerta-roja'>🚨 RUPTURA INMINENTE<br><p style='margin:0; font-size: 14px;'>Impacto a &lt; 7 Días</p><h2 style='margin:0;'>{criticos} Insumos</h2></div>", unsafe_allow_html=True)
                c_k2.markdown(f"<div class='alerta-amarilla'>⚠️ ALERTA LOGÍSTICA<br><p style='margin:0; font-size: 14px;'>Pedir entre 8 y 21 Días</p><h2 style='margin:0;'>{alertas} Insumos</h2></div>", unsafe_allow_html=True)
                c_k3.markdown(f"<div class='alerta-verde'>✅ INVENTARIO SANO<br><p style='margin:0; font-size: 14px;'>Autonomía &gt; 21 Días</p><h2 style='margin:0;'>{optimos} Insumos</h2></div>", unsafe_allow_html=True)
                
                def pintar_oraculo(row):
                    if "CRÍTICO" in row['ESTADO']: return ['background-color: #ffe6e6; color: #cc0000; font-weight:bold; border-bottom:1px solid white;'] * len(row)
                    if "ALERTA" in row['ESTADO']: return ['background-color: #fff3cd; color: #856404; font-weight:bold; border-bottom:1px solid white;'] * len(row)
                    return [''] * len(row)

                df_vista = df_oraculo.copy()
                df_vista['📦 SALDO (SAP)'] = df_vista['📦 SALDO (SAP)'].apply(lambda x: fmt_latino(x, 1))
                df_vista['📈 PROYECCIÓN MES (L/Kg)'] = df_vista['📈 PROYECCIÓN MES (L/Kg)'].apply(lambda x: fmt_latino(x, 1))
                df_vista['⏳ AUTONOMÍA'] = df_vista['⏳ AUTONOMÍA'].apply(lambda x: "∞ (Sin Consumo Histórico)" if x >= 9999 else f"{x:,.0f} Días")

                st.markdown("#### 📋 Detalle de Explosión de Materiales")
                st.dataframe(df_vista.style.apply(pintar_oraculo, axis=1), use_container_width=True, hide_index=True)
                    
            except Exception as e:
                st.error(f"🚨 Falla en los cálculos predictivos o estructura de datos: {e}")
