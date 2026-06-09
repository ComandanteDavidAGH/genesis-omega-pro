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
# 🧠 TRADUCTOR DUAL: CRUCE DE PISTA + EQUIPO
# =================================================================
def asignar_flota_dual(eq_raw, pista_raw):
    eq = str(eq_raw).upper()
    p = str(pista_raw).upper()
    
    if "AEROPENORT" in p:
        if "TRUSH" in eq or "THRUS" in eq: return "THRUS SR2 (AEROPENORT)", 4606562.0
        if "PAWNEE" in eq or "PIPER PA 36" in eq: return "PIPER PA 36-375 (AEROPENORT)", 3985831.0
        if "CESSNA" in eq or "PIPER PA 25" in eq: return "CESSNA O PIPER PA 25 (AEROPENORT)", 3036525.0
        
    if "FUMIGARAY" in p:
        if "AIR TRACTOR" in eq or "TRACTOR" in eq: return "AIR TRACTOR (FUMIGARAY)", 4665107.0
        if "CESSNA" in eq: return "CESSNA FUMIGARAY (FUMIGARAY)", 3065952.0
        
    if "ASA" in p:
        if "CESSNA" in eq: return "CESSNA ASA (ASA)", 3666600.0

    if "DRON" in eq or "DRONE" in eq:
        if "DATAROT" in eq: return "DRONE DATAROT", 84427.0
        if "NORTE" in eq: return "DRONE NORTE", 75518.0
        if "AVIL" in eq: return "DRONE AVIL", 71280.0
        if "GENESYS" in eq: return "DRONE GENESYS", 71280.0
        return f"DRON GENERICO ({p})", 71280.0

    return f"{eq} ({p})", 4606562.0

# =================================================================
# 🚁 MOTOR DEL SIMULADOR SIN TOPES
# =================================================================
def ejecutar(procesar_fecha_pesada, extraer_numero):
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis de Lucro Cesante con Filtros Inteligentes en Cascada.")

    with st.spinner("📥 Sincronizando y cruzando datos de TABLA 1..."):
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
    df_sim = df_sim[df_sim["Equipo"].astype(str).str.strip() != ""]

    df_sim[["Equipo", "Tarifa_Defecto"]] = df_sim.apply(
        lambda r: pd.Series(asignar_flota_dual(r["Equipo"], r["Pista"])), axis=1
    )

    df_sim["Hectareas"] = df_sim["Hectareas"].apply(limpiar_cantidad)
    df_sim["Horometro"] = df_sim["Horometro"].apply(limpiar_cantidad)
    df_sim["CobroReal"] = df_sim["CobroReal"].apply(limpiar_moneda)
    df_sim['Fecha_DT'] = pd.to_datetime(df_sim["Fecha"], dayfirst=True, errors='coerce')
    
    df_sim = df_sim[df_sim["Hectareas"] > 0]

    if df_sim.empty:
        st.warning("📭 No hay registros matemáticamente válidos en la TABLA 1.")
        return

    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime(2023, 1, 1).date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + sorted(df_sim["Finca"].dropna().unique().tolist())
    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + sorted(df_sim["Pista"].dropna().astype(str).unique().tolist())
    
    # Lista maestra global de todos los equipos de toda la empresa (Para inicializar la memoria)
    lista_aviones_maestra = sorted(df_sim["Equipo"].unique().tolist())

    if 'v_maestra_dual' not in st.session_state:
        st.session_state.tarifas_simulador = {}
        dict_temp = dict(zip(df_sim["Equipo"], df_sim["Tarifa_Defecto"]))
        for av in lista_aviones_maestra:
            st.session_state.tarifas_simulador[av] = float(dict_temp.get(av, 4606562.0))
        st.session_state['v_maestra_dual'] = True

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL CON FILTROS EN CASCADA
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario")
        f1, f2, f3, f4, f5, f6 = st.columns([1, 1, 1.2, 1, 1.2, 0.8])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        
        # 🌟 FILTRO EN CASCADA: Extraemos solo la flota de la pista seleccionada
        if pista_sel != "🛣️ TODAS LAS PISTAS":
            lista_aviones_dinamica = sorted(df_sim[df_sim["Pista"] == pista_sel]["Equipo"].unique().tolist())
        else:
            lista_aviones_dinamica = lista_aviones_maestra
            
        opciones_avion_dinamica = ["✈️ TODOS LOS EQUIPOS"] + lista_aviones_dinamica

        equipo_sel = f5.selectbox("✈️ Equipo Fijo", opciones_avion_dinamica)
        multiplicador = f6.number_input("✖️ Mult.", value=1.112, format="%.3f")

        st.markdown("---")
        st.markdown("#### ✈️ Gestor de Tarifas DUAL (Solo equipos de la pista seleccionada)")
        c_tar1, c_tar2 = st.columns(2)
        
        # El gestor también hereda el filtro en cascada para no mezclar aviones
        if not lista_aviones_dinamica:
            lista_aviones_dinamica = ["Sin Equipos"]
            
        avion_editar = c_tar1.selectbox("🚁 Seleccione Aeronave", lista_aviones_dinamica)
        
        tarifa_actual_num = float(st.session_state.tarifas_simulador.get(avion_editar, 0.0))
        tarifa_inicial_formateada = f"$ {tarifa_actual_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        key_dinamica = f"input_dual_{avion_editar.replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '')}"

        tarifa_usuario = c_tar2.text_input(
            f"✍️ Editar Tarifa para {avion_editar} (Presione Enter)", 
            value=tarifa_inicial_formateada,
            key=key_dinamica
        )
        
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

    # --- FILTRAR LOS DATOS PARA LA TABLA MATEMÁTICA ---
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios en las fechas seleccionadas.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO CORREGIDO
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    
    def calcular_costo_ha(row):
        eq = str(row["Equipo"]).upper()
        if "DRON" in eq or "DRONE" in eq or row["Horometro"] == 0:
            return row["Tarifa_Aplicada"] * multiplicador
        else:
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
