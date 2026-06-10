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
def extraer_datos_boveda():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame(), pd.DataFrame()
    try:
        boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        
        t1 = boveda.worksheet("TABLA 1").get_all_values()
        idx_t1 = 4
        for i in range(min(6, len(t1))):
            if "FINCA" in [str(x).upper() for x in t1[i]]:
                idx_t1 = i
                break
        df_t1 = pd.DataFrame(t1[idx_t1+1:], columns=t1[idx_t1]) if len(t1) > idx_t1 else pd.DataFrame()
        
        hojas = [ws.title for ws in boveda.worksheets()]
        nombre_t2 = "TABLA 2" if "TABLA 2" in hojas else hojas[1]
        t2 = boveda.worksheet(nombre_t2).get_all_values()
        df_t2 = pd.DataFrame(t2)
        
        return df_t1, df_t2
    except:
        return pd.DataFrame(), pd.DataFrame()

# 🌟 NUEVOS MOTORES DE LIMPIEZA MATEMÁTICA A PRUEBA DE BALAS
def limpiar_cantidad(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).replace(" ", "").strip()
        if "," in texto and "." in texto:
            if texto.rfind(".") > texto.rfind(","): texto = texto.replace(",", "")
            else: texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except:
        return 0.0

def limpiar_moneda(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        texto = str(val).upper().replace("$", "").replace("COP", "").replace(" ", "").strip()
        if "." in texto and "," in texto:
            if texto.rfind(".") > texto.rfind(","): texto = texto.replace(",", "")
            else: texto = texto.replace(".", "").replace(",", ".")
        else:
            sep = "." if "." in texto else ("," if "," in texto else None)
            if sep:
                if len(texto.split(sep)[-1]) == 3: texto = texto.replace(sep, "")
                else: texto = texto.replace(sep, ".")
        return float(texto)
    except:
        return 0.0

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
# ================================================================
# 🔥 INTEGRACIÓN DE ORDEN DE SERVICIO (Nº ORDEN)
# ================================================================
def integrar_os(df):
    """
    Suma el FactorTiempo (RENDIMIENTO (horas)) por Nº ORDEN
    y devuelve el total a cada fila sin alterar la lógica existente.
    """
    if "Nº ORDEN" not in df.columns:
        st.error("⚠️ Falta la columna 'Nº ORDEN' en TABLA 1.")
        return df

    df_os = df.groupby("Nº ORDEN")["FactorTiempo"].sum().reset_index()
    df_os = df_os.rename(columns={"FactorTiempo": "TiempoTotalOS"})

    df = df.merge(df_os, on="Nº ORDEN", how="left")
    return df

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

def obtener_mult(prod):
    p = str(prod).upper()
    if "TERCERO" in p: return 1.451
    if "AFILIADO" in p: return 1.164
    if "COOPERATIVA" in p: return 1.112
    if "SOCIO" in p: return 1.112
    if "ORGANICO" in p: return 1.011
    return 1.112 

def generar_excel_multi_hoja(df_filtrado_base, df_diario_agrupado, t_real, t_ideal, t_perdido, porcentaje_fuga, titulo_ideal):
    buffer = io.BytesIO()
    
    nombres_meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    df_mes = df_filtrado_base.copy()
    df_mes["Mes_Num"] = df_mes["Fecha_DT"].dt.month.fillna(1).astype(int)
    
    df_mensual_base = df_mes.groupby("Mes_Num").agg({
        "Hectareas": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()
    
    df_mensual_base["Mes de Operación"] = df_mensual_base["Mes_Num"].map(nombres_meses)
    df_mensual_base["Tarifa Real Prom/Ha"] = df_mensual_base["Total Real Facturado"] / df_mensual_base["Hectareas"]
    df_mensual_base["Tarifa Ideal Prom/Ha"] = df_mensual_base["Total Simulado Ideal"] / df_mensual_base["Hectareas"]
    df_mensual_base["Brecha Financiera/Ha"] = df_mensual_base["Tarifa Ideal Prom/Ha"] - df_mensual_base["Tarifa Real Prom/Ha"]
    
    df_mensual_final = df_mensual_base[["Mes de Operación", "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha", "Brecha Financiera/Ha", "Total Real Facturado", "Total Simulado Ideal", "Lucro Cesante"]].copy()
    df_mensual_final = df_mensual_final.rename(columns={"Hectareas": "Total Hectáreas"})

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_mensual_final.to_excel(writer, sheet_name="Resumen_Ejecutivo_Mensual", index=False, startrow=5)
        ws1 = writer.sheets["Resumen_Ejecutivo_Mensual"]
        
        df_diario_renamed = df_diario_agrupado.copy().rename(columns={
            "Hectareas": "Total Ha",
            "Tarifa Real Prom/Ha": "Tarifa Real ($/Ha)",
            "Tarifa Ideal Prom/Ha": f"Tarifa Ideal ({titulo_ideal})",
            "Brecha por Ha": "Brecha ($/Ha)",
            "Total Real Facturado": "Cobro Real Total",
            "Total Simulado Ideal": f"Total Ideal ({titulo_ideal})",
            "Lucro Cesante": "Brecha Financiera Total"
        })
        df_diario_renamed.to_excel(writer, sheet_name="Detalle_Diario_Auditoria", index=False, startrow=5)
        ws2 = writer.sheets["Detalle_Diario_Auditoria"]

        fill_header = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        borde = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws1.cell(row=1, column=1, value="📊 RESUMEN GENERAL DIRECTIVO: CONSOLIDADO MENSUAL").font = Font(size=14, bold=True, color="0D1B2A")
        ws1.cell(row=3, column=1, value=f"💰 Cobro Real Acumulado: $ {t_real:,.0f}").font = Font(bold=True)
        ws1.cell(row=3, column=4, value=f"📈 Ideal ({titulo_ideal}): $ {t_ideal:,.0f}").font = Font(bold=True)
        ws1.cell(row=3, column=7, value=f"⚠️ Fuga Financiera: $ {t_perdido:,.0f} ({porcentaje_fuga:.1f}%)").font = Font(bold=True, color="C00000")

        for col_num in range(1, len(df_mensual_final.columns) + 1):
            cell = ws1.cell(row=6, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws1.column_dimensions[get_column_letter(col_num)].width = 22

        for r in range(7, len(df_mensual_final) + 7):
            ws1.cell(row=r, column=2).number_format = '#,##0.0' 
            for c in range(3, 9): 
                ws1.cell(row=r, column=c).number_format = '"$"#,##0'
            for c in range(1, 9):
                ws1.cell(row=r, column=c).border = borde

        ws2.cell(row=1, column=1, value="📋 INFORME ESPECÍFICO: AUDITORÍA CRONOLÓGICA DIARIA").font = Font(size=14, bold=True, color="0D1B2A")
        
        for col_num in range(1, len(df_diario_renamed.columns) + 1):
            cell = ws2.cell(row=6, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws2.column_dimensions[get_column_letter(col_num)].width = 18

        for r in range(7, len(df_diario_renamed) + 7):
            ws2.cell(row=r, column=5).number_format = '#,##0.0' 
            for c in range(6, 12): 
                ws2.cell(row=r, column=c).number_format = '"$"#,##0'
            for c in range(1, 12):
                ws2.cell(row=r, column=c).border = borde

    return buffer.getvalue()

def ejecutar(procesar_fecha_pesada, extraer_numero):
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Análisis Inteligente: Costo Puro Operativo vs Precio de Venta Ideal.")

    with st.spinner("📥 Sincronizando Inteligencia de TABLA 1 y TABLA 2..."):
        df_base, df_t2_raw = extraer_datos_boveda()

    if df_base.empty:
        st.error("🚨 Base de datos vacía o sin acceso a TABLA 1.")
        return
        
    dict_productores = {}
    if not df_t2_raw.empty:
        idx_t2 = 0
        for i in range(min(5, len(df_t2_raw))):
            if "FINCA" in [str(x).upper() for x in df_t2_raw.iloc[i]]:
                idx_t2 = i; break
        if len(df_t2_raw) > idx_t2 + 1:
            df_t2 = pd.DataFrame(df_t2_raw.values[idx_t2+1:], columns=df_t2_raw.values[idx_t2])
            for idx, row in df_t2.iterrows():
                try:
                    f_name = str(row.iloc[0]).strip().upper()
                    if f_name and f_name != 'NAN':
                        p_tipo = str(row.iloc[5]).strip().upper() if len(row) > 5 else "COOPERATIVA"
                        dict_productores[f_name] = p_tipo
                except: pass

    # Mapeo de columnas según tu TABLA 1
    col_fecha = "FECHA"
    col_finca = "FINCA"
    col_pista = "PISTA"
    col_avion = "MODELO"
    col_ha = "ÁREA FUMIG.\n(ha)"
    col_vuelo = " COSTO AVIÒN\n($/ha) "
    col_orden = "Nº ORDEN"
    col_rend_horas = "RENDIMIENTO (horas)"

    # Detectar columna de tiempo: aquí forzamos a usar RENDIMIENTO (horas)
    if col_rend_horas in df_base.columns:
        col_tiempo = col_rend_horas
    else:
        col_tiempo = None
        cols_upper = {c: str(c).replace("\n", "").strip().upper() for c in df_base.columns}
        for c, c_up in cols_upper.items():
            if "HORO" in c_up: col_tiempo = c; break
        if not col_tiempo:
            for c, c_up in cols_upper.items():
                if "HORAS" in c_up and "HA" not in c_up and "REND" not in c_up: col_tiempo = c; break
        if not col_tiempo:
            for c, c_up in cols_upper.items():
                if "RENDIMIENTO" in c_up or "HORA" in c_up: col_tiempo = c; break
        if not col_tiempo:
            df_base["Factor_Tiempo_Fijo"] = 60.0
            col_tiempo = "Factor_Tiempo_Fijo"

    # Ajuste de columnas requeridas
    for c_req in [col_fecha, col_finca, col_pista, col_avion, col_ha, col_vuelo, col_orden, col_tiempo]:
        if c_req not in df_base.columns:
            posible_match = [c for c in df_base.columns if c_req.replace("\n", "").strip() in c.replace("\n", "").strip()]
            if posible_match:
                if c_req == col_ha: col_ha = posible_match[0]
                if c_req == col_vuelo: col_vuelo = posible_match[0]
                if c_req == col_orden: col_orden = posible_match[0]
                if c_req == col_tiempo: col_tiempo = posible_match[0]

    # Construcción del dataframe base del simulador
    df_sim = df_base[[col_fecha, col_finca, col_pista, col_avion, col_ha, col_tiempo, col_vuelo, col_orden]].copy()
    df_sim.columns = ["Fecha", "Finca", "Pista_Raw", "Equipo_Raw", "Hectareas", "FactorTiempo", "CobroReal", "Nº ORDEN"]
    
    df_sim = df_sim[df_sim["Finca"].astype(str).str.strip() != ""]
    df_sim = df_sim[df_sim["Equipo_Raw"].astype(str).str.strip() != ""]

    df_sim[["Equipo", "Pista"]] = df_sim.apply(lambda r: pd.Series(purificar_datos_vuelo(r["Equipo_Raw"], r["Pista_Raw"])), axis=1)
    df_sim["Hectareas"] = df_sim["Hectareas"].apply(limpiar_cantidad)
    df_sim["FactorTiempo"] = df_sim["FactorTiempo"].apply(limpiar_cantidad)
    df_sim["CobroReal"] = df_sim["CobroReal"].apply(limpiar_moneda)
    df_sim['Fecha_DT'] = df_sim["Fecha"].apply(parsear_fecha_robusta)
    
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

    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + list(FLOTA_OFICIAL_POR_PISTA.keys())
    lista_aviones_maestra = list({
        "THRUS SR2": 4606562.0,
        "PIPER PA 36-375": 3985831.0,
        "CESSNA O PIPER PA 25": 3036525.0,
        "AIR TRACTOR": 4665107.0,
        "CESSNA ASA": 3666600.0,
        "CESSNA FUMIGARAY": 3065952.0,
        "DRONE DATAROT": 84427.0,
        "DRONE NORTE": 75518.0,
        "DRONE AVIL": 71280.0,
        "DRONE GENESYS": 71280.0
    }.keys())

    if 'v_maestra_blindada_11' not in st.session_state:
        st.session_state.tarifas_simulador = {}
        for av in lista_aviones_maestra:
            st.session_state.tarifas_simulador[av] = float({
                "THRUS SR2": 4606562.0,
                "PIPER PA 36-375": 3985831.0,
                "CESSNA O PIPER PA 25": 3036525.0,
                "AIR TRACTOR": 4665107.0,
                "CESSNA ASA": 3666600.0,
                "CESSNA FUMIGARAY": 3065952.0,
                "DRONE DATAROT": 84427.0,
                "DRONE NORTE": 75518.0,
                "DRONE AVIL": 71280.0,
                "DRONE GENESYS": 71280.0
            }.get(av, 4606562.0))
        st.session_state['v_maestra_blindada_11'] = True

    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario")
        f1, f2, f3, f4, f5, f6 = st.columns([1, 1, 1.2, 1, 1.3, 1.5])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        
        if pista_sel != "🛣️ TODAS LAS PISTAS":
            pista_limpia = pista_sel.replace("🛣️ ", "").strip().upper()
            lista_aviones_dinamica = FLOTA_OFICIAL_POR_PISTA.get(pista_limpia, [])
        else:
            lista_aviones_dinamica = lista_aviones_maestra
            
        equipo_sel = f5.selectbox("✈️ Equipo Fijo", ["✈️ TODOS LOS EQUIPOS"] + lista_aviones_dinamica)
        modo_calculo = f6.selectbox("🧮 Analizar Contra:", ["Venta Ideal (+Margen Inteligente)", "Costo Puro Operativo"])

        st.markdown("---")
        st.markdown(f"#### ✈️ Gestor de Tarifas Base ({pista_sel.replace('🛣️ ', '')})")
        
        equipos_a_mostrar = [av for av in lista_aviones_dinamica if av != "✈️ TODOS LOS EQUIPOS"]
        if not equipos_a_mostrar:
            st.info("📭 Seleccione una pista válida para ver y editar su flota oficial.")
        else:
            for avion_editar in equipos_a_mostrar:
                c_nombre, c_precio = st.columns([1.5, 2])
                c_nombre.markdown(
                    f"<div style='margin-top: 5px; font-weight: bold; color: #1a365d; font-size: 15px;'>🚁 {avion_editar}</div>",
                    unsafe_allow_html=True
                )
                tarifa_actual_num = float(st.session_state.tarifas_simulador.get(avion_editar, 0.0))
                tarifa_inicial_formateada = f"$ {tarifa_actual_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                
                tarifa_usuario = c_precio.text_input(
                    "Tarifa",
                    value=tarifa_inicial_formateada,
                    key=f"in_bl_{avion_editar.replace(' ', '_')}",
                    label_visibility="collapsed"
                )
                
                if tarifa_usuario != tarifa_inicial_formateada:
                    try:
                        limpio = tarifa_usuario.replace("$", "").replace(" ", "").strip()
                        if "," in limpio and "." in limpio: limpio = limpio.replace(".", "").replace(",", ".")
                        elif "." in limpio and len(limpio.split(".")[-1]) == 3: limpio = limpio.replace(".", "")
                        elif "," in limpio: limpio = limpio.replace(",", ".")
                        st.session_state.tarifas_simulador[avion_editar] = float(limpio)
                        st.rerun()
                    except: pass

        tarifas_aviones = st.session_state.tarifas_simulador

    # ============================
    # 🔍 FILTRADO DE ESCENARIO
    # ============================
    df_filtrado = df_sim[
        (df_sim["Fecha_DT"].dt.date >= fecha_ini) &
        (df_sim["Fecha_DT"].dt.date <= fecha_fin)
    ].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS":
        df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS":
        df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel.replace("🛣️ ", "")]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS":
        df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios en las fechas seleccionadas.")
        return

    # ============================
    # ============================
    # ============================
    # 🔥 INTEGRACIÓN ORDEN DE SERVICIO
    # ============================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    df_filtrado["Fecha Operación"] = df_filtrado["Fecha_DT"].dt.strftime("%Y-%m-%d")
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]

    # Limpiar merges anteriores si existen
    for col_extra in ["TiempoTotalOS", "HectareasTotalOS", "TiempoTotalOS_x", 
                      "TiempoTotalOS_y", "HectareasTotalOS_x", "HectareasTotalOS_y"]:
        if col_extra in df_filtrado.columns:
            df_filtrado = df_filtrado.drop(columns=[col_extra])

    # 1. Agrupar por Nº ORDEN
    df_os = df_filtrado.groupby("Nº ORDEN").agg({
        "FactorTiempo": "sum",
        "Hectareas": "sum"
    }).reset_index().rename(columns={
        "FactorTiempo": "TiempoTotalOS",
        "Hectareas": "HectareasTotalOS"
    })

    # 2. Unir al detalle
    df_filtrado = df_filtrado.merge(df_os, on="Nº ORDEN", how="left")

    # 3. Verificar que las columnas existen antes de calcular
    if "TiempoTotalOS" not in df_filtrado.columns or "HectareasTotalOS" not in df_filtrado.columns:
        st.error("⚠️ Error: No se pudo calcular TiempoTotalOS. Verifica la columna 'Nº ORDEN'.")
        return

    # 4. Fórmula SIN TOPES
    def calcular_ideal_por_os(row):
        try:
            valor_hora = float(row["Tarifa_Aplicada"]) if pd.notna(row["Tarifa_Aplicada"]) else 0.0
            tiempo_total = float(row["TiempoTotalOS"]) if pd.notna(row["TiempoTotalOS"]) else 0.0
            ha_total = float(row["HectareasTotalOS"]) if pd.notna(row["HectareasTotalOS"]) else 0.0
            if ha_total == 0:
                return 0.0
            return (valor_hora * tiempo_total) / ha_total
        except:
            return 0.0

    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(calcular_ideal_por_os, axis=1)
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]
    # ============================
    # 📊 AGRUPACIÓN RESUMEN
    # ============================
    df_agrupado = df_filtrado.groupby(["Fecha Operación", "Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()
    
    df_agrupado["Tarifa Real Prom/Ha"] = df_agrupado["Total Real Facturado"] / df_agrupado["Hectareas"]
    df_agrupado["Tarifa Ideal Prom/Ha"] = df_agrupado["Total Simulado Ideal"] / df_agrupado["Hectareas"]
    df_agrupado["Brecha por Ha"] = df_agrupado["Tarifa Ideal Prom/Ha"] - df_agrupado["Tarifa Real Prom/Ha"]

    df_agrupado = df_agrupado[[
        "Fecha Operación", "Pista", "Finca", "Equipo",
        "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha",
        "Brecha por Ha", "Total Real Facturado", "Total Simulado Ideal", "Lucro Cesante"
    ]]
    df_agrupado = df_agrupado.sort_values(by=["Finca", "Fecha Operación"])

    # ============================
    # 📊 TABLA EN PANTALLA
    # ============================
    st.markdown("### 📊 Resumen por Día / Finca / Equipo")
    st.dataframe(
        df_agrupado.style.format({
            "Hectareas": "{:,.2f}",
            "Tarifa Real Prom/Ha": "{:,.0f}",
            "Tarifa Ideal Prom/Ha": "{:,.0f}",
            "Brecha por Ha": "{:,.0f}",
            "Total Real Facturado": "{:,.0f}",
            "Total Simulado Ideal": "{:,.0f}",
            "Lucro Cesante": "{:,.0f}"
        }),
        use_container_width=True,
        height=420
    )

    # ============================
    # 📈 GRÁFICOS
    # ============================
    st.markdown("### 📈 Comparativo de Tarifas Promedio por Hectárea")

    fig_tarifas = px.bar(
        df_agrupado,
        x="Fecha Operación",
        y=["Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha"],
        color_discrete_sequence=["#1f77b4", "#ff7f0e"],
        barmode="group",
        hover_data=["Finca", "Equipo", "Hectareas"],
        labels={"value": "Tarifa ($/ha)", "Fecha Operación": "Fecha"},
        title="Tarifa Real vs Tarifa Ideal por Día"
    )
    st.plotly_chart(fig_tarifas, use_container_width=True)

    st.markdown("### 💰 Lucro Cesante por Día")

    df_lucro = df_agrupado.groupby("Fecha Operación")["Lucro Cesante"].sum().reset_index()
    fig_lucro = px.bar(
        df_lucro,
        x="Fecha Operación",
        y="Lucro Cesante",
        labels={"Lucro Cesante": "Lucro Cesante ($)", "Fecha Operación": "Fecha"},
        color_discrete_sequence=["#d62728"],
        title="Lucro Cesante Total por Día"
    )
    st.plotly_chart(fig_lucro, use_container_width=True)

    # ============================
    # 📤 EXPORTAR A EXCEL
    # ============================
    st.markdown("### 📤 Exportar Escenario a Excel")

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import io

    def construir_excel(df_detalle, df_resumen):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "DETALLE VUELOS"

        cols_det = list(df_detalle.columns)
        for j, col in enumerate(cols_det, start=1):
            cell = ws1.cell(row=1, column=j, value=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1f4e78")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for i, row in enumerate(df_detalle.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                ws1.cell(row=i, column=j, value=val)

        for j in range(1, len(cols_det) + 1):
            ws1.column_dimensions[get_column_letter(j)].width = 18

        ws2 = wb.create_sheet(title="RESUMEN AGRUPADO")
        cols_res = list(df_resumen.columns)
        for j, col in enumerate(cols_res, start=1):
            cell = ws2.cell(row=1, column=j, value=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1f4e78")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for i, row in enumerate(df_resumen.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                ws2.cell(row=i, column=j, value=val)

        for j in range(1, len(cols_res) + 1):
            ws2.column_dimensions[get_column_letter(j)].width = 20

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    df_detalle_export = df_filtrado.copy()
    df_detalle_export = df_detalle_export[[
        "Fecha Operación", "Nº ORDEN", "Finca", "Pista", "Equipo",
        "Hectareas", "FactorTiempo", "TiempoTotalOS",
        "Tarifa_Aplicada", "CobroReal",
        "Costo Simulado HA", "Total Real Facturado", "Total Simulado Ideal", "Lucro Cesante"
    ]]

    buffer_excel = construir_excel(df_detalle_export, df_agrupado)

    st.download_button(
        label="📥 Descargar Escenario en Excel",
        data=buffer_excel,
        file_name=f"Simulador_Financiero_Libre_{fecha_ini}_{fecha_fin}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Escenario calculado con Orden de Servicio integrada (Nº ORDEN + RENDIMIENTO (horas)).")

if __name__ == "__main__":
    pass
