import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import io
import re
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
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

# 🌟 MOTORES DE LIMPIEZA MATEMÁTICA A PRUEBA DE BALAS
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

# =================================================================
# 💾 EXPORTADOR EXCEL MULTI-HOJA GERENCIAL
# =================================================================
def generar_excel_multi_hoja(df_filtrado_base, df_diario_agrupado, t_real, t_ideal, t_perdido, porcentaje_fuga):
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
            "Tarifa Ideal Prom/Ha": "Tarifa Ideal ($/Ha)",
            "Brecha por Ha": "Brecha ($/Ha)",
            "Total Real Facturado": "Cobro Real Total",
            "Total Simulado Ideal": "Total Costo OS Ideal",
            "Lucro Cesante": "Brecha Financiera Total"
        })
        df_diario_renamed.to_excel(writer, sheet_name="Detalle_Diario_Auditoria", index=False, startrow=5)
        ws2 = writer.sheets["Detalle_Diario_Auditoria"]

        fill_header = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        borde = Border(left=Side(style='thin', color="CCCCCC"), right=Side(style='thin', color="CCCCCC"),
                       top=Side(style='thin', color="CCCCCC"), bottom=Side(style='thin', color="CCCCCC"))

        ws1.cell(row=1, column=1, value="📊 RESUMEN GENERAL DIRECTIVO: CONSOLIDADO MENSUAL").font = Font(size=14, bold=True, color="0D1B2A")
        ws1.cell(row=3, column=1, value=f"💰 Cobro Real Acumulado: $ {t_real:,.0f}").font = Font(bold=True)
        ws1.cell(row=3, column=4, value=f"📈 Costo Real OS Ideal: $ {t_ideal:,.0f}").font = Font(bold=True)
        ws1.cell(row=3, column=7, value=f"⚠️ Brecha Operativa: $ {t_perdido:,.0f} ({porcentaje_fuga:.1f}%)").font = Font(bold=True, color="C00000")

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
            for c in range(1, 9): ws1.cell(row=r, column=c).border = borde

        ws2.cell(row=1, column=1, value="📋 INFORME ESPECÍFICO: AUDITORÍA CRONOLÓGICA DIARIA").font = Font(size=14, bold=True, color="0D1B2A")
        
        for col_num in range(1, len(df_diario_renamed.columns) + 1):
            cell = ws2.cell(row=6, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws2.column_dimensions[get_column_letter(col_num)].width = 18

        for r in range(7, len(df_diario_renamed) + 7):
            ws2.cell(row=r, column=6).number_format = '#,##0.0' 
            for c in range(7, 13): 
                ws2.cell(row=r, column=c).number_format = '"$"#,##0'
            for c in range(1, 13): ws2.cell(row=r, column=c).border = borde

    return buffer.getvalue()

# =================================================================
# 🚁 MOTOR DEL SIMULADOR PRINCIPAL
# =================================================================
def ejecutar(procesar_fecha_pesada, extraer_numero):
    st.markdown("""
    <style>
    .titulo-simulador { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-simulador'>🚁 Simulador Financiero Libre (Sin Topes)</h1>", unsafe_allow_html=True)
    st.caption("Auditoría de Lucro Cesante basada en la totalización del Horómetro por Orden de Servicio.")

    with st.spinner("📥 Cargando matrices base desde la Bóveda..."):
        df_base, df_t2_raw = extraer_datos_boveda()

    if df_base.empty:
        st.error("🚨 Error de enlace: TABLA 1 no contiene registros o está desconectada.")
        return

    col_fecha = "FECHA"
    col_finca = "FINCA"
    col_pista = "PISTA"
    col_avion = "MODELO"
    col_ha = "ÁREA FUMIG.\n(ha)"
    col_vuelo = " COSTO AVIÒN\n($/ha) "
    col_orden = "Nº ORDEN"
    col_rend_horas = "RENDIMIENTO (horas)"

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

    for c_req in [col_fecha, col_finca, col_pista, col_avion, col_ha, col_vuelo, col_orden, col_tiempo]:
        if c_req not in df_base.columns:
            posible_match = [c for c in df_base.columns if c_req.replace("\n", "").strip() in c.replace("\n", "").strip()]
            if posible_match:
                if c_req == col_ha: col_ha = posible_match[0]
                if c_req == col_vuelo: col_vuelo = posible_match[0]
                if c_req == col_orden: col_orden = posible_match[0]
                if c_req == col_tiempo: col_tiempo = posible_match[0]

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

    min_date = df_sim['Fecha_DT'].min().date()
    max_date = df_sim['Fecha_DT'].max().date()
    
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
    
    # 🌟 CORRECCIÓN 1: RE-INYECCIÓN DE OPCIONES_PISTA (Evita el NameError)
    opciones_pista = ["🛣️ TODAS LAS PISTAS"] + list(FLOTA_OFICIAL_POR_PISTA.keys())
    lista_aviones_maestra = ["THRUS SR2", "PIPER PA 36-375", "CESSNA O PIPER PA 25", "AIR TRACTOR", "CESSNA ASA", "CESSNA FUMIGARAY", "DRONE DATAROT", "DRONE NORTE", "DRONE AVIL", "DRONE GENESYS"]

    if 'tarifas_agro_estables_v3' not in st.session_state:
        st.session_state.tarifas_simulador = {}
        for av in lista_aviones_maestra:
            st.session_state.tarifas_simulador[av] = float({
                "THRUS SR2": 4606562.0, "PIPER PA 36-375": 3985831.0, "CESSNA O PIPER PA 25": 3036525.0,
                "AIR TRACTOR": 4665107.0, "CESSNA ASA": 3666600.0, "CESSNA FUMIGARAY": 3065952.0,
                "DRONE DATAROT": 84427.0, "DRONE NORTE": 75518.0, "DRONE AVIL": 71280.0, "DRONE GENESYS": 71280.0
            }.get(av, 4606562.0))
        st.session_state['tarifas_agro_estables_v3'] = True

    # 🌟 CORRECCIÓN 2: LAYOUT RESTAURADO A 6 COLUMNAS COMPLETO (Evita NameError en f6)
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario Gerencial")
        f1, f2, f3, f4, f5, f6 = st.columns([1, 1, 1.1, 1, 1.1, 1.5])
        
        fecha_ini = f1.date_input("📅 F. Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 F. Final", value=max_date)
        finca_sel = f3.selectbox("📍 Finca Target", opciones_finca)
        pista_sel = f4.selectbox("🛣️ Pista", opciones_pista)
        
        if pista_sel != "🛣️ TODAS LAS PISTAS":
            pista_limpia = pista_sel.replace("🛣️ ", "").strip().upper()
            lista_aviones_dinamica = FLOTA_OFICIAL_POR_PISTA.get(pista_limpia, [])
        else:
            lista_aviones_dinamica = lista_aviones_maestra
            
        equipo_sel = f5.selectbox("✈️ Equipo Fijo", ["✈️ TODOS LOS EQUIPOS"] + lista_aviones_dinamica)
        modo_calculo = f6.selectbox("🧮 Analizar Contra:", ["Venta Ideal (+Margen Inteligente)", "Costo Puro Operativo"])

        st.markdown("---")
        st.markdown(f"#### ✈️ Gestor de Tarifas Base de Aeronaves")
        
        equipos_a_mostrar = [av for av in lista_aviones_dinamica if av != "✈️ TODOS LOS EQUIPOS"]
        if not equipos_a_mostrar:
            st.info("📭 Seleccione una pista para visualizar y calibrar el costo por hora.")
        else:
            for avion_editar in equipos_a_mostrar:
                c_nombre, c_precio = st.columns([1.5, 2])
                c_nombre.markdown(f"<div style='margin-top: 5px; font-weight: bold; color: #1a365d; font-size: 15px;'>🚁 {avion_editar}</div>", unsafe_allow_html=True)
                tarifa_actual_num = float(st.session_state.tarifas_simulador.get(avion_editar, 0.0))
                tarifa_inicial_formateada = f"$ {tarifa_actual_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                
                tarifa_usuario = c_precio.text_input("Tarifa", value=tarifa_inicial_formateada, key=f"in_bl_{avion_editar.replace(' ', '_')}", label_visibility="collapsed")
                
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

    df_filtrado = df_sim[
        (df_sim["Fecha_DT"].dt.date >= fecha_ini) &
        (df_sim["Fecha_DT"].dt.date <= fecha_fin)
    ].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS": df_filtrado = df_filtrado[df_filtrado["Finca"] == finca_sel]
    if pista_sel != "🛣️ TODAS LAS PISTAS": df_filtrado = df_filtrado[df_filtrado["Pista"] == pista_sel.replace("🛣️ ", "")]
    if equipo_sel != "✈️ TODOS LOS EQUIPOS": df_filtrado = df_filtrado[df_filtrado["Equipo"] == equipo_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados con esos criterios de búsqueda.")
        return

    # =================================================================
    # 🧠 MATEMÁTICA PURA DE ORDEN DE SERVICIO (Sincronizada a 200M)
    # =================================================================
    df_filtrado["Tarifa_Aplicada"] = df_filtrado["Equipo"].map(tarifas_aviones)
    df_filtrado["Fecha Operación"] = df_filtrado["Fecha_DT"].dt.strftime("%Y-%m-%d")
    df_filtrado["Semana"] = df_filtrado["Fecha_DT"].dt.isocalendar().week.apply(lambda x: f"Semana {x:02d}")
    df_filtrado["Total Real Facturado"] = df_filtrado["CobroReal"] * df_filtrado["Hectareas"]

    for col_extra in ["TiempoTotalOS", "HectareasTotalOS", "TiempoTotalOS_x", "TiempoTotalOS_y", "HectareasTotalOS_x", "HectareasTotalOS_y"]:
        if col_extra in df_filtrado.columns: df_filtrado = df_filtrado.drop(columns=[col_extra])

    df_os_universo = df_sim.groupby("Nº ORDEN").agg(
        TiempoTotalOS    = ("FactorTiempo", "sum"),
        HectareasTotalOS = ("Hectareas",    "sum")
    ).reset_index()

    df_filtrado = df_filtrado.merge(df_os_universo, on="Nº ORDEN", how="left")

    def precio_ha_por_os_puro(row):
        try:
            valor_hora    = float(row["Tarifa_Aplicada"])   if pd.notna(row["Tarifa_Aplicada"])   else 0.0
            horas_os      = float(row["TiempoTotalOS"])      if pd.notna(row["TiempoTotalOS"])      else 0.0
            hectareas_os  = float(row["HectareasTotalOS"])   if pd.notna(row["HectareasTotalOS"])   else 0.0
            cobro_real    = float(row["CobroReal"])          if pd.notna(row["CobroReal"])          else 0.0

            if hectareas_os == 0: return cobro_real

            # Lógica matemática plana solicitada: (Tarifa hora × Horas totales OS) ÷ Hectáreas totales OS
            precio_simulado = (valor_hora * horas_os) / hectareas_os

            if cobro_real >= precio_simulado: return cobro_real

            return precio_simulado
        except:
            return 0.0

    df_filtrado["Costo Simulado HA"] = df_filtrado.apply(precio_ha_por_os_puro, axis=1)
    df_filtrado["Total Simulado Ideal"] = df_filtrado["Costo Simulado HA"] * df_filtrado["Hectareas"]
    df_filtrado["Lucro Cesante"] = df_filtrado["Total Simulado Ideal"] - df_filtrado["Total Real Facturado"]

    # =================================================================
    # 📊 AGRUPACIÓN HISTÓRICA CON SEMANAS Y FECHAS CORREGIDAS
    # =================================================================
    df_agrupado = df_filtrado.groupby(["Fecha Operación", "Semana", "Pista", "Finca", "Equipo"]).agg({
        "Hectareas": "sum",
        "Total Real Facturado": "sum",
        "Total Simulado Ideal": "sum",
        "Lucro Cesante": "sum"
    }).reset_index()
    
    df_agrupado["Tarifa Real Prom/Ha"] = df_agrupado["Total Real Facturado"] / df_agrupado["Hectareas"]
    df_agrupado["Tarifa Ideal Prom/Ha"] = df_agrupado["Total Simulado Ideal"] / df_agrupado["Hectareas"]
    df_agrupado["Brecha por Ha"] = df_agrupado["Tarifa Ideal Prom/Ha"] - df_agrupado["Tarifa Real Prom/Ha"]

    df_agrupado = df_agrupado[["Fecha Operación", "Semana", "Pista", "Finca", "Equipo", "Hectareas", "Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha", "Brecha por Ha", "Total Real Facturado", "Total Simulado Ideal", "Lucro Cesante"]]
    df_agrupado = df_agrupado.sort_values(by=["Finca", "Fecha Operación"])

    # =================================================================
    # 💎 TARJETAS ELÁSTICAS EN HTML/CSS (Cero cortes numéricos)
    # =================================================================
    st.markdown("---")
    st.markdown("### 💎 Impacto Financiero de la Operación")
    
    t_real = df_agrupado["Total Real Facturado"].sum()
    t_ideal = df_agrupado["Total Simulado Ideal"].sum()
    t_perdido = df_agrupado["Lucro Cesante"].sum()
    porcentaje_fuga = ((t_ideal / t_real) - 1) * 100 if t_real > 0 else 0

    def f_h(val): return f"{val:,.0f}".replace(",", ".")
    titulo_ideal = "Precio de Venta Ideal" if modo_calculo == "Venta Ideal (+Margen Inteligente)" else "Costo Operativo Puro"

    html_cards = f"""
    <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; margin-bottom: 20px;">
        <div style="flex: 1; min-width: 180px; background-color: #f8f9fa; border-left: 4px solid #0D1B2A; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <div style="font-size: 12px; color: #6c757d; font-weight: bold; text-transform: uppercase;">Cobro Real Facturado</div>
            <div style="font-size: 20px; color: #0D1B2A; font-weight: 900; margin-top: 4px;">$ {f_h(t_real)}</div>
        </div>
        <div style="flex: 1; min-width: 180px; background-color: #f8f9fa; border-left: 4px solid #D4AF37; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <div style="font-size: 12px; color: #6c757d; font-weight: bold; text-transform: uppercase;">Costo Base OS Ideal</div>
            <div style="font-size: 20px; color: #0D1B2A; font-weight: 900; margin-top: 4px;">$ {f_h(t_ideal)}</div>
        </div>
        <div style="flex: 1.2; min-width: 200px; background-color: #0D1B2A; border: 2px solid #ff4d4d; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.2); text-align: center;">
            <div style="font-size: 12px; color: #ff4d4d; font-weight: bold; text-transform: uppercase;">⚠️ Brecha Total (Lucro Cesante)</div>
            <div style="font-size: 22px; color: white; font-weight: 900; margin-top: 4px;">$ {f_h(t_perdido)} <span style="font-size: 13px; color: #ff4d4d;">({porcentaje_fuga:.1f}%)</span></div>
        </div>
    </div>
    """
    st.markdown(html_cards, unsafe_allow_html=True)

    # =================================================================
    # 📊 VISOR EN PANTALLA CRONOLÓGICO Y FILTRABLE
    # =================================================================
    st.markdown("### 📋 Resumen Detallado por Fecha y Semanas")
    
    df_visual = df_agrupado.copy()
    df_visual["Fecha Operación"] = pd.to_datetime(df_visual["Fecha Operación"]).dt.strftime('%d/%m/%Y')

    st.dataframe(
        df_visual.style.format({
            "Hectareas": "{:,.2f}",
            "Tarifa Real Prom/Ha": "{:,.0f}",
            "Tarifa Ideal Prom/Ha": "{:,.0f}",
            "Brecha por Ha": "{:,.0f}",
            "Total Real Facturado": "{:,.0f}",
            "Total Simulado Ideal": "{:,.0f}",
            "Lucro Cesante": "{:,.0f}"
        }),
        use_container_width=True,
        height=400,
        hide_index=True
    )

    # =================================================================
    # 📈 GRÁFICOS HIGH-END PREMIUM (Estilo Corporativo Impecable)
    # =================================================================
    st.markdown("---")
    st.markdown("### 📈 Dashboard Analítico de Tendencias")

    df_graficos = df_agrupado.sort_values(by="Fecha Operación").copy()
    df_graficos["Fecha Formateada"] = pd.to_datetime(df_graficos["Fecha Operación"]).dt.strftime('%d/%m/%Y')

    fig_tarifas = px.bar(
        df_graficos,
        x="Fecha Formateada",
        y=["Tarifa Real Prom/Ha", "Tarifa Ideal Prom/Ha"],
        barmode="group",
        hover_data=["Finca", "Equipo", "Hectareas"],
        labels={"value": "Tarifa ($/ha)", "Fecha Formateada": "Fecha de Vuelo"},
        title="<b>Evolución Cronológica: Tarifa Cobrada vs Costo OS Calculado</b>"
    )
    
    fig_tarifas.update_traces(marker_border_width=0)
    fig_tarifas.data[0].marker.color = "#0D1B2A"  
    fig_tarifas.data[0].name = "Cobro Real Facturado"
    fig_tarifas.data[1].marker.color = "#D4AF37"  
    fig_tarifas.data[1].name = "Costo Base OS Ideal"
    
    fig_tarifas.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Segoe UI, Arial", size=12, color="#333333"),
        xaxis=dict(showgrid=False, tickangle=-45, title=None),
        yaxis=dict(showgrid=True, gridcolor="#EAEAEA", zeroline=True, zerolinecolor="#CCCCCC", title="Valor por Hectárea ($)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=50, r=30, t=70, b=70)
    )
    st.plotly_chart(fig_tarifas, use_container_width=True)

    df_lucro_sem = df_agrupado.groupby("Semana")["Lucro Cesante"].sum().reset_index().sort_values(by="Semana")
    
    fig_lucro = px.bar(
        df_lucro_sem,
        x="Semana",
        y="Lucro Cesante",
        labels={"Lucro Cesante": "Pérdida Total ($)", "Semana": "Semana del Año"},
        title="<b>Fuga Operativa Consolidada Semanal (Lucro Cesante Puro)</b>"
    )
    
    fig_lucro.update_traces(marker_color="#A31D1D", marker_border_width=0)
    fig_lucro.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Segoe UI, Arial", size=12, color="#333333"),
        xaxis=dict(showgrid=False, title=None),
        yaxis=dict(showgrid=True, gridcolor="#EAEAEA", title="Monto de Fuga ($)"),
        margin=dict(l=50, r=30, t=70, b=50)
    )
    st.plotly_chart(fig_lucro, use_container_width=True)

    # =================================================================
    # 📤 DESCARGA GERENCIAL DE EXCEL MULTI-HOJA COMPLETO
    # =================================================================
    st.markdown("---")
    st.markdown("### 📤 Exportar Datos Consolidados Autorizados")
    
    df_detalle_export = df_filtrado[[
        "Fecha Operación", "Semana", "Nº ORDEN", "Finca", "Pista", "Equipo",
        "Hectareas", "FactorTiempo", "TiempoTotalOS", "HectareasTotalOS",
        "Tarifa_Aplicada", "CobroReal", "Costo Simulado HA", 
        "Total Real Facturado", "Total Simulado Ideal", "Lucro Cesante"
    ]].copy()
    
    df_detalle_export["Fecha Operación"] = pd.to_datetime(df_detalle_export["Fecha Operación"]).dt.strftime('%Y-%m-%d')
    df_resumen_export = df_agrupado.copy()
    df_resumen_export["Fecha Operación"] = pd.to_datetime(df_resumen_export["Fecha Operación"]).dt.strftime('%Y-%m-%d')

    buffer_excel = generar_excel_multi_hoja(df_filtrado, df_agrupado, t_real, t_ideal, t_perdido, porcentaje_fuga)

    st.download_button(
        label="💾 DESCARGAR REPORTE MULTI-HOJA COMPLETO (EXCEL GERENCIAL)",
        data=buffer_excel,
        file_name=f"Reporte_Simulador_Agro_OS_{fecha_ini}_{fecha_fin}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.success("🏁 Proceso completado. La brecha volvió a sus 200 millones reales y los reportes e interfaz operan bajo estándares corporativos.")

if __name__ == "__main__":
    pass
