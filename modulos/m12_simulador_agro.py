import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 CONEXIÓN A BÓVEDA DE DATOS
# =================================================================
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

@st.cache_data(show_spinner=False, ttl=600)
def extraer_tabla_1_historica():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        t1 = boveda.worksheet("TABLA 1").get_all_values()
        idx_t1 = 4
        for i in range(min(6, len(t1))):
            if "FINCA" in [str(x).upper() for x in t1[i]]:
                idx_t1 = i
                break
        
        if len(t1) > idx_t1:
            df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1])
            df_t1 = df_t1.loc[:, ~df_t1.columns.duplicated()]
            return df_t1
        return pd.DataFrame()
    except:
        return pd.DataFrame()

def limpiar_cantidad(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).replace(" ", "").strip()
        if "," in texto: texto = texto.replace(",", ".")
        return float(texto)
    except:
        return 0.0

def limpiar_moneda(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).upper().replace("$", "").replace("COP", "").replace(" ", "").strip()
        if "." in texto:
            partes = texto.split(".")
            if len(partes[-1]) == 3: texto = texto.replace(".", "")
        return float(texto)
    except:
        return 0.0

# =================================================================
# 🚁 MOTOR DEL SIMULADOR SIN TOPES EN BASE A TU ESTRUCTURA REAL
# =================================================================
def ejecutar(procesar_fecha_pesada, extraer_numero):
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis de Lucro Cesante por Finca, Pista y Tipo de Aeronave real.")

    with st.spinner("📥 Sincronizando historial de operaciones (TABLA 1)..."):
        df_base = extraer_tabla_1_historica()

    if df_base.empty:
        st.error("🚨 Base de datos vacía o sin acceso a TABLA 1.")
        return

    col_fecha = "FECHA"
    col_finca = "FINCA"
    col_pista = "PISTA"
    col_avion = "MODELO"
    col_ha = "ÁREA FUMIG.\n(ha)"
    col_horo = "RENDIMIENTO (horas)" 
    col_vuelo = "COSTO AVIÒN\n($/ha)"

    for c_req in [col_fecha, col_finca, col_pista, col_avion, col_ha, col_horo, col_vuelo]:
        if c_req not in df_base.columns:
            posible_match = [c for c in df_base.columns if c_req.replace("\n", "").strip() in c.replace("\n", "").strip()]
            if posible_match:
                if c_req == col_ha: col_ha = posible_match[0]
                if c_req == col_vuelo: col_vuelo = posible_match[0]
                if c_req == col_horo: col_horo = posible_match[0]

    df_sim = df_base[[col_fecha, col_finca, col_pista, col_avion, col_ha, col_horo, col_vuelo]].copy()
    df_sim.columns = ["Fecha", "Finca", "Pista", "Equipo", "Hectareas", "Horometro", "CobroReal"]
    
    df_sim = df_sim[df_sim["Finca"].astype(str).str.strip() != ""] 
    df_sim["Equipo"] = df_sim["Equipo"].astype(str).str.strip().str.upper()
    df_sim = df_sim[df_sim["Equipo"] != ""]

    df_sim["Hectareas"] = df_sim["Hectareas"].apply(limpiar_cantidad)
    df_sim["Horometro"] = df_sim["Horometro"].apply(limpiar_cantidad)
    df_sim["CobroReal"] = df_sim["CobroReal"].apply(limpiar_moneda)
    df_sim['Fecha_DT'] = pd.to_datetime(df_sim["Fecha"], dayfirst=True, errors='coerce')
    
    df_sim = df_sim[df_sim["Hectareas"] > 0]

    if df_sim.empty:
        st.warning("📭 No hay registros matemáticamente válidos (con hectáreas > 0) in la TABLA 1.")
        return

    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime(2023, 1, 1).date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + sorted(df_sim["Finca"].dropna().unique().tolist())
    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + sorted(df_sim["Pista"].dropna().astype(str).unique().tolist())
    opciones_avion = ["✈️ TODOS LOS EQUIPOS"] + sorted(df_sim["Equipo"].dropna().astype(str).unique().tolist())
    lista_aviones_pura = sorted(df_sim["Equipo"].dropna().astype(str).unique().tolist())

    # =================================================================
    # 🧠 MAESTRO DE TARIFAS REALES (Extraídas de tu pestaña Validación Dosis)
    # =================================================================
    if 'tarifas_simulador' not in st.session_state:
        st.session_state.tarifas_simulador = {}
        
        # Diccionario con valores exactos de tu captura de pantalla
        tarifas_oficiales_hoja = {
            "THRUS SR2": 4606562.0,
            "PIPER PA 36-375": 3985831.0,
            "CESSNA O PIPER PA 25": 3036525.0,
            "AIR TRACTOR": 4665107.0,
            "CESSNA ASA": 3666600.0,
            "CESSNA FUMIGARAY": 3065952.0,
            "DRONE DATAROT": 84427.0,
            "DRONE GENESYS": 71280.0,
            "DRONE AVIL": 71280.0,
            "DRONE NORTE": 75518.0
        }
        
        for avion in lista_aviones_pura:
            avion_upper = avion.upper()
            val_asignado = 4606562.0  # Por defecto si no se encuentra match
            
            for nombre_ref, precio in tarifas_oficiales_hoja.items():
                if nombre_ref in avion_upper or avion_upper in nombre_ref:
                    val_asignado = precio
                    break
            
            st.session_state.tarifas_simulador[avion] = float(val_asignado)

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL 
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario")
        f1, f2, f3, f4, f5, f6 = st.columns([1, 1, 1.2, 1, 1.2, 0.8])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        equipo_sel = f5.selectbox("✈️ Equipo Fijo", opciones_avion)
        multiplicador = f6.number_input("✖️ Mult.", value=1.112, format="%.3f")

        st.markdown("---")
        st.markdown("#### ✈️ Gestor de Tarifas de Flota")
        c_tar1, c_tar2 = st.columns(2)
        
        # CASILLA 1: Elegir Avión
        avion_editar = c_tar1.selectbox("🚁 Seleccione Aeronave a configurar", lista_aviones_pura)
        
        # Obtener el valor numérico actual desde la memoria reactiva
        tarifa_actual_num = float(st.session_state.tarifas_simulador.get(avion_editar, 0.0))
        
        # Formatear el número crudo a estética ejecutiva colombiana ($ 4.665.107,00)
        tarifa_inicial_formateada = f"$ {tarifa_actual_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Generamos una llave dinámica única basada en el avión para obligar el cambio inmediato en la UI
        key_dinamica = f"input_tar_din_{avion_editar.replace(' ', '_').replace('-', '_')}"

        # CASILLA 2: Cuadro de texto reactivo alineado al modelo de avión
        tarifa_usuario = c_tar2.text_input(
            f"✍️ Editar Tarifa para {avion_editar} (Presione Enter para aplicar)", 
            value=tarifa_inicial_formateada,
            key=key_dinamica
        )
        
        # 🧠 PROCESAMIENTO INTELIGENTE: Si cambias el número y das Enter, se sobrescribe en caliente
        if tarifa_usuario != tarifa_inicial_formateada:
            try:
                limpio = tarifa_usuario.replace("$", "").replace(" ", "").strip()
                if "," in limpio and "." in limpio:
                    limpio = limpio.replace(".", "").replace(",", ".")
                elif "." in limpio and len(limpio.split(".")[-1]) == 3:
                    limpio = limpio.replace(".", "")
                elif "," in limpio:
                    limpio = limpio.replace(",", ".")
                    
                valor_final_numerico = float(limpio)
                st.session_state.tarifas_simulador[avion_editar] = valor_final_numerico
                st.rerun()
            except:
                pass

        tarifas_aviones = st.session_state.tarifas_simulador

    # --- FILTRAR LOS DATOS ---
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios en las fechas seleccionadas.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO CORREGIDO (Diferencia si es Dron o Avión)
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    
    def calcular_costo_ha(row):
        eq = str(row["Equipo"]).upper()
        # Si es dron, la tarifa de la hoja de cálculo ya es por HECTÁREA de forma directa
        if "DRON" in eq or "DRONE" in eq or row["Horometro"] == 0:
            return row["Tarifa_Aplicada"] * multiplicador
        else:
            # Si es avión, la tarifa es por HORA, por ende se calcula usando el horómetro
            return ((row["Tarifa_Aplicada"] * row["Horometro"]) / row["Hectareas"]) * multiplicador

    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(calcular_costo_ha, axis=1)
    
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    df_agrupado = df_filtrado.groupby(["Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
        "Horometro": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()

    df_agrupado["Tarifa Real Prom/Ha"] = df_agrupado["Total Real Facturado"] / df_agrupado["Hectareas"]
    df_agrupado["Tarifa Ideal Prom/Ha"] = df_agrupado["Total Simulado Ideal"] / df_agrupado["Hectareas"]
    df_agrupado["Brecha por Ha"] = df_agrupado["Tarifa Ideal Prom/Ha"] - df_agrupado["Tarifa Real Prom/Ha"]

    # =================================================================
    # 📊 DASHBOARD DE MÉTRICAS EJECUTIVAS
    # =================================================================
    st.markdown("---")
    st.markdown("### 💎 Impacto Financiero de la Operación")
    
    t_real = df_agrupado["Total Real Facturado"].sum()
    t_ideal = df_agrupado["Total Simulado Ideal"].sum()
    t_perdido = df_agrupado["Lucro Cesante"].sum()
    porcentaje_fuga = ((t_ideal / t_real) - 1) * 100 if t_real > 0 else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Cobro Real Registrado (Con Topes)", f"$ {t_real:,.0f}".replace(",", "."))
    m2.metric("Cobro Matemático Puro (Sin Topes)", f"$ {t_ideal:,.0f}".replace(",", "."))
    m3.metric("⚠️ Lucro Cesante (Brecha Total)", f"$ {t_perdido:,.0f}".replace(",", "."), delta=f"{porcentaje_fuga:.1f}% de fuga", delta_color="inverse")

    st.markdown("#### 📉 Comparativa Facturación Total por Finca")
    df_g_resumen = df_agrupado.groupby("Finca")[["Total Real Facturado", "Total Simulado Ideal"]].sum().reset_index()
    df_g = df_g_resumen.melt(id_vars="Finca", value_vars=["Total Real Facturado", "Total Simulado Ideal"], var_name="Escenario", value_name="Monto ($)")
    fig = px.bar(df_g, x="Finca", y="Monto ($)", color="Escenario", barmode="group",
                 color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Ideal": "#d4af37"})
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, legend=dict(yanchor="top", y=1.1, xanchor="left", x=0.01))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 📋 Análisis Detallado: Brecha por Hectárea y Total")
    df_mostrar = df_agrupado[["Pista", "Finca", "Equipo", "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha", "Brecha por Ha", "Lucro Cesante"]].copy()
    
    df_mostrar["Hectareas"] = df_mostrar["Hectareas"].apply(lambda x: f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", "."))
    df_mostrar["Tarifa Real Prom/Ha"] = df_mostrar["Tarifa Real Prom/Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Tarifa Ideal Prom/Ha"] = df_mostrar["Tarifa Ideal Prom/Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Brecha por Ha"] = df_mostrar["Brecha por Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Lucro Cesante"] = df_mostrar["Lucro Cesante"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    
    st.dataframe(df_mostrar.sort_values(by=["Finca"]), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    pass
