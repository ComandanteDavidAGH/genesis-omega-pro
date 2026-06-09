import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import io
import re
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# 🌟 LECTOR ROBUSTO DE FECHAS (Caza fantasmas como "Enero")
def parsear_fecha_robusta(val):
    if pd.isna(val) or str(val).strip() == "": return pd.NaT
    s = str(val).strip().lower()
    if s.isdigit(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(int(s), 'D')
    
    meses = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12}
    match1 = re.search(r'(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})', s)
    if match1:
        dia_str, mes_str, anio_str = match1.groups()
        if mes_str in meses: return pd.to_datetime(f"{anio_str}-{meses[mes_str]:02d}-{int(dia_str):02d}")
    match2 = re.search(r'([a-z]+)\s+(\d{1,2}),\s+(\d{4})', s)
    if match2:
        mes_str, dia_str, anio_str = match2.groups()
        if mes_str in meses: return pd.to_datetime(f"{anio_str}-{meses[mes_str]:02d}-{int(dia_str):02d}")
    try: 
        return pd.to_datetime(s.split(" ")[0], dayfirst=True, errors='coerce')
    except: 
        return pd.NaT

# =================================================================
# 🧠 TRADUCTOR BLINDADO
# =================================================================
def purificar_datos_vuelo(eq_raw, pista_raw):
    eq = str(eq_raw).upper()
    p = str(pista_raw).upper()
    
    if "DRON" in eq or "DRONE" in eq:
        if "DATAROT" in eq or "PLUC" in p: return "DRONE DATAROT", "PLUC"
        if "NORTE" in eq or "PDIV" in p: return "DRONE NORTE", "PDIV"
        if "AVIL" in eq or "TEHO" in p: return "DRONE AVIL", "TEHO"
        if "GENESYS" in eq or "LUCI" in p: return "DRONE GENESYS", "LUCI"
        return "DRONE GENESYS", "LUCI" 
        
    if "TRUSH" in eq or "THRUS" in eq or "OMANDER" in eq: return "THRUS SR2", "AEROPENORT"
    if "PAWNEE" in eq or "BRAVO" in eq or "PIPER PA 36" in eq: return "PIPER PA 36-375", "AEROPENORT"
    if "AIR TRACTOR" in eq or "TRACTOR" in eq or "TOR" in eq: return "AIR TRACTOR", "FUMIGARAY"
    
    if "CESSNA" in eq or "PIPER PA 25" in eq:
        if "ASA" in p or "ASA" in eq: return "CESSNA ASA", "ASA"
        if "FUMIGARAY" in p or "FUMIGARAY" in eq: return "CESSNA FUMIGARAY", "FUMIGARAY"
        return "CESSNA O PIPER PA 25", "AEROPENORT"

    return "IGNORAR", "IGNORAR"

# =================================================================
# 💾 EXPORTADOR EXCEL PROFESIONAL
# =================================================================
def generar_excel_profesional(df_agrupado, t_real, t_ideal, t_perdido, porcentaje_fuga):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_ex = df_agrupado.copy()
        df_ex = df_ex.rename(columns={
            "Hectareas": "Total Ha",
            "Tarifa Real Prom/Ha": "Tarifa Real ($/Ha)",
            "Tarifa Ideal Prom/Ha": "Tarifa Ideal ($/Ha)",
            "Brecha por Ha": "Brecha ($/Ha)",
            "Total Real Facturado": "Cobro Real Total",
            "Total Simulado Ideal": "Costo Operativo Ideal",
            "Lucro Cesante": "Brecha Financiera Total"
        })
        
        if "FactorTiempo" in df_ex.columns: df_ex = df_ex.drop(columns=["FactorTiempo"])
        
        df_ex.to_excel(writer, sheet_name="Resumen_Financiero", index=False, startrow=5)
        ws = writer.sheets["Resumen_Financiero"]

        fill_header = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        borde = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws.cell(row=1, column=1, value="REPORTE DE IMPACTO FINANCIERO Y COSTOS OPERATIVOS").font = Font(size=14, bold=True, color="0D1B2A")
        ws.cell(row=3, column=1, value=f"💰 Cobro Real: $ {t_real:,.0f}").font = Font(bold=True)
        ws.cell(row=3, column=4, value=f"📈 Costo Ideal Pleno: $ {t_ideal:,.0f}").font = Font(bold=True)
        ws.cell(row=3, column=7, value=f"⚠️ Brecha de Fuga: $ {t_perdido:,.0f} ({porcentaje_fuga:.1f}%)").font = Font(bold=True, color="C00000")

        for col_num, col_name in enumerate(df_ex.columns, 1):
            cell = ws.cell(row=6, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 18

        for r in range(7, len(df_ex) + 7):
            ws.cell(row=r, column=4).number_format = '#,##0.0' # Ha
            for c in range(5, 11): # Dinero
                ws.cell(row=r, column=c).number_format = '"$"#,##0'
            for c in range(1, 11):
                ws.cell(row=r, column=c).border = borde

    return buffer.getvalue()

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
    st.caption("Análisis de Costos Planos 100% Puros (Márgenes Desactivados).")

    with st.spinner("📥 Sincronizando y depurando datos de TABLA 1..."):
        df_base = extraer_tabla_1_historica()

    if df_base.empty:
        st.error("🚨 Base de datos vacía o sin acceso a TABLA 1.")
        return

    col_fecha, col_finca, col_pista, col_avion, col_ha, col_vuelo = "FECHA", "FINCA", "PISTA", "MODELO", "ÁREA FUMIG.\n(ha)", "COSTO AVIÒN\n($/ha)"

    col_tiempo = None
    for c in df_base.columns:
        c_up = str(c).upper()
        if "RENDIMIENTO" in c_up or "HORA" in c_up or "HORO" in c_up:
            col_tiempo = c
            break
            
    if not col_tiempo:
        df_base["Factor_Tiempo"] = 60.0
        col_tiempo = "Factor_Tiempo"

    for c_req in [col_fecha, col_finca, col_pista, col_avion, col_ha, col_vuelo]:
        if c_req not in df_base.columns:
            posible_match = [c for c in df_base.columns if c_req.replace("\n", "").strip() in c.replace("\n", "").strip()]
            if posible_match:
                if c_req == col_ha: col_ha = posible_match[0]
                if c_req == col_vuelo: col_vuelo = posible_match[0]

    df_sim = df_base[[col_fecha, col_finca, col_pista, col_avion, col_ha, col_tiempo, col_vuelo]].copy()
    df_sim.columns = ["Fecha", "Finca", "Pista_Raw", "Equipo_Raw", "Hectareas", "FactorTiempo", "CobroReal"]
    
    df_sim = df_sim[df_sim["Finca"].astype(str).str.strip() != ""] 
    df_sim = df_sim[df_sim["Equipo_Raw"].astype(str).str.strip() != ""]

    df_sim[["Equipo", "Pista"]] = df_sim.apply(
        lambda r: pd.Series(purificar_datos_vuelo(r["Equipo_Raw"], r["Pista_Raw"])), axis=1
    )

    df_sim["Hectareas"] = df_sim["Hectareas"].apply(limpiar_cantidad)
    df_sim["FactorTiempo"] = df_sim["FactorTiempo"].apply(limpiar_cantidad)
    df_sim["CobroReal"] = df_sim["CobroReal"].apply(limpiar_moneda)
    
    # 🌟 APLICAMOS LA IA CAZAFANTASMAS DE FECHAS
    df_sim['Fecha_DT'] = df_sim["Fecha"].apply(parsear_fecha_robusta)
    
    # Excluimos errores crudos y datos nulos
    df_sim = df_sim[(df_sim["Hectareas"] > 0) & (df_sim["Equipo"] != "IGNORAR") & (df_sim['Fecha_DT'].notna())]

    if df_sim.empty:
        st.warning("📭 No hay registros matemáticamente válidos en la TABLA 1.")
        return

    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime(2023, 1, 1).date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + sorted(df_sim["Finca"].dropna().unique().tolist())
    
    FLOTA_OFICIAL_POR_PISTA = {
        "AEROPENORT": ["THRUS SR2", "PIPER PA 36-375", "CESSNA O PIPER PA 25"],
        "FUMIGARAY": ["AIR TRACTOR", "CESSNA FUMIGARAY"],
        "ASA": ["CESSNA ASA"],
        "PLUC": ["DRONE DATAROT"],
        "PDIV": ["DRONE NORTE"],
        "TEHO": ["DRONE AVIL"],
        "LUCI": ["DRONE GENESYS"]
    }
    
    PRECIOS_OFICIALES = {
        "THRUS SR2": 4606562.0, "PIPER PA 36-375": 3985831.0, "CESSNA O PIPER PA 25": 3036525.0,
        "AIR TRACTOR": 4665107.0, "CESSNA ASA": 3666600.0, "CESSNA FUMIGARAY": 3065952.0,
        "DRONE DATAROT": 84427.0, "DRONE NORTE": 75518.0, "DRONE AVIL": 71280.0, "DRONE GENESYS": 71280.0
    }

    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + list(FLOTA_OFICIAL_POR_PISTA.keys())
    lista_aviones_maestra = list(PRECIOS_OFICIALES.keys())

    if 'v_maestra_blindada_7' not in st.session_state:
        st.session_state.tarifas_simulador = {}
        for av in lista_aviones_maestra:
            st.session_state.tarifas_simulador[av] = float(PRECIOS_OFICIALES.get(av, 4606562.0))
        st.session_state['v_maestra_blindada_7'] = True

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL (Interfaz Limpia, sin Multiplicadores)
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario")
        f1, f2, f3, f4, f5 = st.columns([1, 1, 1.5, 1, 1.5])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        
        if pista_sel != "🛣️ TODAS LAS PISTAS":
            pista_limpia = pista_sel.replace("🛣️ ", "").strip().upper()
            lista_aviones_dinamica = FLOTA_OFICIAL_POR_PISTA.get(pista_limpia, [])
        else:
            lista_aviones_dinamica = lista_aviones_maestra
            
        opciones_avion_dinamica = ["✈️ TODOS LOS EQUIPOS"] + lista_aviones_dinamica
        equipo_sel = f5.selectbox("✈️ Equipo Fijo", opciones_avion_dinamica)

        st.markdown("---")
        st.markdown(f"#### ✈️ Gestor de Tarifas ({pista_sel.replace('🛣️ ', '')})")
        
        equipos_a_mostrar = [av for av in lista_aviones_dinamica if av != "✈️ TODOS LOS EQUIPOS"]
        
        if not equipos_a_mostrar:
            st.info("📭 Seleccione una pista válida para ver y editar su flota oficial.")
        else:
            for avion_editar in equipos_a_mostrar:
                c_nombre, c_precio = st.columns([1.5, 2])
                
                c_nombre.markdown(f"<div style='margin-top: 5px; font-weight: bold; color: #1a365d; font-size: 15px;'>🚁 {avion_editar}</div>", unsafe_allow_html=True)
                
                tarifa_actual_num = float(st.session_state.tarifas_simulador.get(avion_editar, 0.0))
                tarifa_inicial_formateada = f"$ {tarifa_actual_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                key_dinamica = f"in_blind7_{avion_editar.replace(' ', '_').replace('-', '_')}"
                
                tarifa_usuario = c_precio.text_input(
                    "Tarifa Base Plana", 
                    value=tarifa_inicial_formateada,
                    key=key_dinamica,
                    label_visibility="collapsed" 
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

    # --- FILTRAR LOS DATOS ---
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel.replace("🛣️ ", "")]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios en las fechas seleccionadas.")
        return

    # =================================================================
    # 🧠 MOTOR FINANCIERO IA (100% PLANO)
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    
    def calcular_costo_ha_plano(row):
        eq = str(row["Equipo"]).upper()
        tarifa = float(row["Tarifa_Aplicada"])
        val_tiempo = float(row["FactorTiempo"])
        ha = float(row["Hectareas"])
        
        if "DRON" in eq or "DRONE" in eq:
            return tarifa 
            
        if val_tiempo == 0 or ha == 0:
            return tarifa / 60.0 

        if val_tiempo > 15: # Rendimiento (Ha/Hr)
            return tarifa / val_tiempo 
        else: # Horómetro (Horas)
            velocidad_implicada = ha / val_tiempo
            if velocidad_implicada > 150:
                return tarifa / 60.0 
            else:
                return (tarifa * val_tiempo) / ha 

    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(calcular_costo_ha_plano, axis=1)
    
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    # 🌟 AGRUPACIÓN SIMPLE Y LIMPIA
    df_agrupado = df_filtrado.groupby(["Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
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

    def f_h(val): return f"{val:,.0f}".replace(",", ".")

    html_cards = f"""
    <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; margin-bottom: 20px;">
        <div style="flex: 1; min-width: 180px; background-color: #f8f9fa; border-left: 4px solid #1b263b; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <div style="font-size: 12px; color: #6c757d; font-weight: 800; text-transform: uppercase;">Cobro Real Histórico</div>
            <div style="font-size: 20px; color: #0d1b2a; font-weight: 900; margin-top: 4px;">$ {f_h(t_real)}</div>
        </div>
        <div style="flex: 1; min-width: 180px; background-color: #f8f9fa; border-left: 4px solid #d4af37; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <div style="font-size: 12px; color: #6c757d; font-weight: 800; text-transform: uppercase;">Costo Plano Puro (Sin Márgenes)</div>
            <div style="font-size: 20px; color: #0d1b2a; font-weight: 900; margin-top: 4px;">$ {f_h(t_ideal)}</div>
        </div>
        <div style="flex: 1.2; min-width: 200px; background-color: #0d1b2a; border: 2px solid #ff4d4d; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.2); text-align: center;">
            <div style="font-size: 12px; color: #ff4d4d; font-weight: 800; text-transform: uppercase;">⚠️ Brecha Financiera Operativa</div>
            <div style="font-size: 22px; color: white; font-weight: 900; margin-top: 4px;">$ {f_h(t_perdido)}</div>
        </div>
    </div>
    """
    st.markdown(html_cards, unsafe_allow_html=True)

    st.markdown("#### 📉 Comparativa Costo Plano vs Facturación Real por Finca")
    df_g_resumen = df_agrupado.groupby("Finca")[["Total Real Facturado", "Total Simulado Ideal"]].sum().reset_index()
    df_g = df_g_resumen.melt(id_vars="Finca", value_vars=["Total Real Facturado", "Total Simulado Ideal"], var_name="Escenario", value_name="Monto ($)")
    fig = px.bar(df_g, x="Finca", y="Monto ($)", color="Escenario", barmode="group",
                 color_discrete_map={"Total Real Facturado": "#1b263b", "Total Simulado Ideal": "#d4af37"})
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, legend=dict(yanchor="top", y=1.1, xanchor="left", x=0.01))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 📋 Análisis Detallado: Costo Puro por Hectárea")
    
    df_mostrar = df_agrupado[["Pista", "Finca", "Equipo", "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha", "Brecha por Ha", "Lucro Cesante"]].copy()
    
    df_mostrar["Hectareas"] = df_mostrar["Hectareas"].apply(lambda x: f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", "."))
    df_mostrar["Tarifa Real Prom/Ha"] = df_mostrar["Tarifa Real Prom/Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Tarifa Ideal Prom/Ha"] = df_mostrar["Tarifa Ideal Prom/Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Brecha por Ha"] = df_mostrar["Brecha por Ha"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    df_mostrar["Lucro Cesante"] = df_mostrar["Lucro Cesante"].apply(lambda x: f"$ {x:,.0f}".replace(",", "."))
    
    st.dataframe(df_mostrar.sort_values(by=["Finca"]), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📑 Reportes Gerenciales")
    st.download_button(
        label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL PROFESIONAL)",
        data=generar_excel_profesional(df_agrupado, t_real, t_ideal, t_perdido, porcentaje_fuga),
        file_name=f"Simulador_Financiero_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if __name__ == "__main__":
    pass
