import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
from oauth2client.service_account import ServiceAccountCredentials

# --- CONEXIÓN Y UTILIDADES ---
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

# --- CEREBRO DEL ORÁCULO ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-oraculo { color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }
    .alerta-roja { background-color: #ffe6e6; color: #cc0000; padding: 10px; border-left: 5px solid #cc0000; border-radius: 5px; font-weight: bold; margin-bottom: 10px;}
    .alerta-amarilla { background-color: #fff3cd; color: #856404; padding: 10px; border-left: 5px solid #ffc107; border-radius: 5px; font-weight: bold; margin-bottom: 10px;}
    .alerta-verde { background-color: #d4edda; color: #155724; padding: 10px; border-left: 5px solid #28a745; border-radius: 5px; font-weight: bold; margin-bottom: 10px;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-oraculo'>🔮 Módulo 13: El Oráculo (Proyección de Autonomía)</h1>", unsafe_allow_html=True)
    st.write("Cruce predictivo entre Existencias Físicas (SAP) y Tasa de Consumo Operativo Reciente.")

    # 1. CARGA DEL INVENTARIO ACTUAL
    st.markdown("### 📥 1. Radar de Existencias Actuales")
    archivo_sap = st.file_uploader("Cargue la Sábana SAP actualizada (.xlsx o .csv)", type=['xlsx', 'csv'], key="sap_oraculo")
    
    if not archivo_sap:
        st.info("💡 Esperando la Sábana SAP para iniciar la predicción de rupturas de stock.")
        return

    with st.spinner("Despertando al Oráculo y calculando tasas de consumo..."):
        try:
            # Leer SAP
            if archivo_sap.name.lower().endswith('.xlsx') or archivo_sap.name.lower().endswith('.xls'):
                df_sap = pd.read_excel(archivo_sap)
            else:
                try: df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='utf-8')
                except:
                    archivo_sap.seek(0)
                    df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='latin1')

            cols = [str(c).upper().strip() for c in df_sap.columns]
            df_sap.columns = cols
            
            c_prod = next((c for c in cols if 'DESC' in c or 'PRODUCTO' in c or 'TEXTO' in c), None)
            c_pista = next((c for c in cols if 'ALMACEN' in c or 'PISTA' in c), None)
            c_saldo = next((c for c in cols if 'LIBRE' in c or 'SALDO' in c), None)

            if not c_prod or not c_saldo:
                st.error("❌ No se detectaron columnas válidas en el archivo SAP.")
                return

            df_sap['SALDO_FISICO'] = df_sap[c_saldo].apply(a_numero_limpio)
            df_sap_agrupado = df_sap.groupby([c_pista, c_prod])['SALDO_FISICO'].sum().reset_index() if c_pista else df_sap.groupby(c_prod)['SALDO_FISICO'].sum().reset_index()
            df_sap_agrupado = df_sap_agrupado[df_sap_agrupado['SALDO_FISICO'] > 0]

            # 2. CALCULAR TASA DE QUEMA (Últimos 30 Días de TABLA 1)
            gc = inicializar_cliente_gspread()
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            
            t1_data = boveda.worksheet("TABLA 1").get_all_values()
            df_t1 = pd.DataFrame(t1_data[5:], columns=[str(c).upper().strip() for c in t1_data[4]])
            
            mezclas_data = boveda.worksheet("DD_Mesclas").get_all_values()
            df_mezclas = pd.DataFrame(mezclas_data[1:], columns=[str(c).upper().strip() for c in mezclas_data[0]])

            # Filtrar TABLA 1 a últimos 30 días
            df_t1['FECHA_DT'] = pd.to_datetime(df_t1['FECHA'], dayfirst=True, errors='coerce')
            hace_30_dias = datetime.now() - timedelta(days=30)
            df_t1_reciente = df_t1[df_t1['FECHA_DT'] >= hace_30_dias].copy()

            df_t1_reciente['HA_NETAS'] = df_t1_reciente['AREA_FUMIG.'].apply(a_numero_limpio)
            ha_por_coctel = df_t1_reciente.groupby('COCTEL')['HA_NETAS'].sum().reset_index()

            # Explotar Recetas
            consumo_30d = {}
            for _, row_c in ha_por_coctel.iterrows():
                coctel_nombre = str(row_c['COCTEL']).upper().strip().split(" ")[0] # Base del cóctel
                ha_aplicadas = row_c['HA_NETAS']
                
                receta = df_mezclas[df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip() == coctel_nombre]
                for _, r_mez in receta.iterrows():
                    prod_quimico = str(r_mez.iloc[1]).upper().strip()
                    dosis = a_numero_limpio(r_mez.iloc[2])
                    if prod_quimico not in ["NAN", "", "NONE"] and dosis > 0:
                        consumo_30d[prod_quimico] = consumo_30d.get(prod_quimico, 0) + (dosis * ha_aplicadas)

            # 3. CRUCE DEL ORÁCULO
            resultados = []
            for _, row_s in df_sap_agrupado.iterrows():
                pista = row_s[c_pista] if c_pista else "GENERAL"
                producto_sap = str(row_s[c_prod]).upper().strip()
                saldo = row_s['SALDO_FISICO']

                # Buscar coincidencia de consumo
                consumo_mes = 0.0
                for p_receta, vol_mes in consumo_30d.items():
                    if p_receta in producto_sap or producto_sap in p_receta:
                        consumo_mes += vol_mes

                consumo_diario = consumo_mes / 30 if consumo_mes > 0 else 0
                dias_autonomia = saldo / consumo_diario if consumo_diario > 0 else 999

                if consumo_diario > 0: # Solo analizamos lo que realmente se está moviendo
                    if dias_autonomia <= 7:
                        estado, color = "🚨 CRÍTICO (< 7 Días)", "#ffe6e6"
                    elif dias_autonomia <= 15:
                        estado, color = "⚠️ ALERTA (8-15 Días)", "#fff3cd"
                    else:
                        estado, color = "✅ ÓPTIMO (> 15 Días)", "#d4edda"

                    resultados.append({
                        "PISTA": pista,
                        "PRODUCTO": producto_sap,
                        "SALDO ACTUAL (L/Kg)": saldo,
                        "CONSUMO DIARIO (L/Kg)": round(consumo_diario, 1),
                        "DÍAS DE AUTONOMÍA": round(dias_autonomia, 0),
                        "ESTADO": estado
                    })

            df_oraculo = pd.DataFrame(resultados).sort_values(by="DÍAS DE AUTONOMÍA")

            st.markdown("---")
            st.markdown("### 🎯 Tablero de Autonomía de Hangares")
            
            if df_oraculo.empty:
                st.success("No se detectaron movimientos recientes que amenacen los inventarios ingresados.")
            else:
                criticos = len(df_oraculo[df_oraculo['DÍAS DE AUTONOMÍA'] <= 7])
                alertas = len(df_oraculo[(df_oraculo['DÍAS DE AUTONOMÍA'] > 7) & (df_oraculo['DÍAS DE AUTONOMÍA'] <= 15)])
                
                c_k1, c_k2, c_k3 = st.columns(3)
                c_k1.markdown(f"<div class='alerta-roja'>🚨 RUPTURA INMINENTE<br><h2 style='margin:0;'>{criticos} Insumos</h2></div>", unsafe_allow_html=True)
                c_k2.markdown(f"<div class='alerta-amarilla'>⚠️ PRONÓSTICO DE COMPRA<br><h2 style='margin:0;'>{alertas} Insumos</h2></div>", unsafe_allow_html=True)
                
                def pintar_oraculo(row):
                    if "CRÍTICO" in row['ESTADO']: return ['background-color: #ffe6e6; color: #cc0000; font-weight:bold;'] * len(row)
                    if "ALERTA" in row['ESTADO']: return ['background-color: #fff3cd; color: #856404; font-weight:bold;'] * len(row)
                    return [''] * len(row)

                st.dataframe(
                    df_oraculo.style.apply(pintar_oraculo, axis=1).format({
                        "SALDO ACTUAL (L/Kg)": "{:,.1f}",
                        "CONSUMO DIARIO (L/Kg)": "{:,.1f}",
                        "DÍAS DE AUTONOMÍA": "{:,.0f} Días"
                    }), 
                    use_container_width=True, hide_index=True
                )
                
        except Exception as e:
            st.error(f"Falla en los cálculos predictivos: {e}")
