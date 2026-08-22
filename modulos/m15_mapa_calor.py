import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import io
import requests
import folium
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =================================================================
# ⚙️ CONSTANTES CENTRALIZADAS (ÚNICA FUENTE DE VERDAD)
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

# =================================================================
# ⚡ MOTOR DE CONEXIÓN UNIFICADO (V42 VIP)
# =================================================================
@st.cache_resource(show_spinner=False)
def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    if "gcp_credentials" in st.secrets:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_credentials"]), scope)
            return gspread.authorize(creds)
        except Exception: pass
    try:
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

# =================================================================
# 🛡️ UTILIDADES DE PURIFICACIÓN Y KML
# =================================================================
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

def limpiar_nombre(texto):
    txt = re.sub(r'[^\w]', '', str(texto).upper())
    txt = txt.replace("FINCA", "").replace("KML", "")
    return txt

def extraer_poligonos_kml(kml_bytes):
    try:
        texto = kml_bytes.decode("utf-8", errors="ignore")
        bloques = re.findall(r'<coordinates>(.*?)</coordinates>', texto, re.IGNORECASE | re.DOTALL)
        poligonos_finca = []
        for bloque in bloques:
            coordenadas_crudas = bloque.strip().split()
            puntos = []
            for coord in coordenadas_crudas:
                partes = coord.split(',')
                if len(partes) >= 2:
                    try:
                        lon = float(partes[0].strip())
                        lat = float(partes[1].strip())
                        puntos.append([lat, lon]) 
                    except: pass
            if len(puntos) >= 3: 
                poligonos_finca.append(puntos)
        return poligonos_finca
    except: return []

# =================================================================
# 🛰️ CONEXIÓN SATELITAL CLIMÁTICA
# =================================================================
@st.cache_data(show_spinner=False, ttl=3600)
def consultar_clima_satelital(lat, lon):
    try:
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&past_days=90&daily=precipitation_sum&timezone=America/Bogota"
        res = requests.get(url, timeout=10).json()
        
        if "daily" in res and "precipitation_sum" in res["daily"]:
            df_clima = pd.DataFrame({
                'fecha': pd.to_datetime(res['daily']['time']),
                'lluvia': [x if x is not None else 0.0 for x in res['daily']['precipitation_sum']]
            })
            
            hoy = pd.to_datetime(datetime.now().date())
            hace_30_dias = hoy - pd.Timedelta(days=30)
            
            lluvia_90d = df_clima[df_clima['fecha'] <= hoy]['lluvia'].sum()
            lluvia_30d = df_clima[(df_clima['fecha'] <= hoy) & (df_clima['fecha'] >= hace_30_dias)]['lluvia'].sum()
            
            return lluvia_90d, lluvia_30d
    except Exception: pass
    return 0.0, 0.0

# =================================================================
# 💾 EXTRACCIÓN CACHEADA
# =================================================================
@st.cache_data(show_spinner=False, ttl=3600)
def cargar_historico_t1():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame()
    
    try:
        boveda = gc.open_by_url(URL_BOVEDA_MAESTRA)
        t1_vals = boveda.worksheet("TABLA 1").get_all_values()
        
        idx_t1 = 4
        for i in range(min(8, len(t1_vals))):
            fila_limpia = [str(x).upper().strip() for x in t1_vals[i]]
            if "Nº ORDEN" in fila_limpia or "FINCA" in fila_limpia:
                idx_t1 = i
                break
                
        df_t1 = pd.DataFrame(t1_vals[idx_t1+1:], columns=[str(c).upper().strip() for c in t1_vals[idx_t1]])
        
        col_fecha = next((c for c in df_t1.columns if 'FECHA' in c), 'FECHA')
        col_ha = next((c for c in df_t1.columns if 'NETA' in c or 'FUMIG' in c or 'HECT' in c), None)
        col_sector = next((c for c in df_t1.columns if 'SECTOR' in c), 'SECTOR')
        col_finca = next((c for c in df_t1.columns if 'FINCA' in c), 'FINCA')
        
        if col_fecha and col_ha:
            df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
            df_t1 = df_t1.dropna(subset=['FECHA_DT'])
            df_t1['HA_CALCULO'] = df_t1[col_ha].apply(a_numero_limpio)
            df_t1['SECTOR_NOM'] = df_t1[col_sector].astype(str).str.upper().str.strip()
            df_t1['FINCA_NOM'] = df_t1[col_finca].astype(str).str.upper().str.strip()
            return df_t1
    except: pass
    return pd.DataFrame()

# =================================================================
# 📤 EXPORTADOR EXCEL VIP
# =================================================================
def generar_excel_vip(df, sheet_name="Mapa_Agronomico"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        header_font = Font(color="D4AF37", bold=True)
        borde_fino = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), 
                            top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
                            
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = borde_fino
            
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
                
                if cell.row > 1:
                    cell.border = borde_fino
                    if isinstance(cell.value, (int, float)):
                        col_name = str(ws[col_letter + '1'].value).upper()
                        if "LLUVIA" in col_name or "DÍAS" in col_name or "RETORNO" in col_name:
                            cell.number_format = '#,##0.0'
                            
            ws.column_dimensions[col_letter].width = min(max_length + 4, 30)
            
    return buffer.getvalue()

# =================================================================
# 🚀 EJECUCIÓN PRINCIPAL
# =================================================================
def ejecutar(purificar_lote, extraer_numero):
    VERDE_INTENSO = '#143521'
    DORADO = '#d4af37'

    st.markdown(f"""
    <style>
    .titulo-agronomo {{ color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }}
    
    [data-testid="column"] {{
        display: flex !important;
        flex-direction: column !important;
        justify-content: flex-start !important;
        align-items: stretch !important;
    }}

    div[data-testid="stDataFrame"] {{ border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.08) !important; }}
    
    div[data-testid="stFileUploader"] > div {{
        background-color: #ffffff !important;
        border: 2px solid {VERDE_INTENSO} !important;
        border-radius: 8px !important;
        box-shadow: 0px 4px 8px rgba(0,0,0,0.06) !important;
    }}
    div[data-testid="stFileUploader"] * {{
        color: #000000 !important;
        font-weight: bold !important;
    }}
    div[data-testid="stFileUploader"] button, div[data-testid="stFileUploader"] button * {{
        color: #ffffff !important;
    }}
    div[data-testid="stMainBlockContainer"] label p {{
        color: #0d1b2a !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
    }}
    </style>
    """, unsafe_allow_html=True)

    def tarjeta_kpi(titulo, valor, delta_texto="", color_delta="#28a745"):
        delta_html = f"<span style='font-size: 14px; color: {color_delta}; margin-left: 8px; vertical-align: middle; padding: 2px 6px; border-radius: 4px; background-color: rgba(255,255,255,0.1);'>{delta_texto}</span>" if delta_texto else ""
        return f"""
        <div style='background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #27AE60; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 20px; height: 100%; min-height: 85px; display: flex; flex-direction: column; justify-content: center;'>
            <p style='font-size: 11px; font-weight: bold; color: #27AE60; text-transform: uppercase; margin:0 0 5px 0; letter-spacing: 1px;'>{titulo}</p>
            <p style='font-size: 22px; font-family: "Arial Black", sans-serif; margin: 0; color: white; display: flex; align-items: center;'>{valor} {delta_html}</p>
        </div>
        """

    c_tit, c_sync = st.columns([3.5, 1.5])
    with c_tit:
        st.markdown("<h1 class='titulo-agronomo'>🗺️ Módulo 15: Mapa de Calor Agronómico</h1>", unsafe_allow_html=True)
        st.write("Análisis de ciclos biológicos por FINCA sobre terreno satelital y lluvia trimestral.")
    with c_sync:
        st.write("")
        if st.button("🔄 Sincronizar Nube (Forzar Datos)", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.rerun()

    with st.container(border=True):
        st.markdown("### 📂 1. Inyección de Polígonos de Precisión")
        
        if "kml_reset_key" not in st.session_state:
            st.session_state.kml_reset_key = 0

        c_file, c_btn_trash = st.columns([3, 1])
        archivos_kml = c_file.file_uploader(
            "Arrastre aquí los archivos .kml de sus fincas", 
            type=['kml'], 
            accept_multiple_files=True, 
            key=f"kml_uploader_{st.session_state.kml_reset_key}",
            label_visibility="collapsed"
        )

        with c_btn_trash:
            if st.button("🗑️ Vaciar Bandeja de KMLs", type="secondary", use_container_width=True):
                st.session_state.kml_reset_key += 1
                st.rerun()

        coor_estimadas = {
            "ORIHUECA": [10.7483, -74.1542], "FLORIDA": [10.7650, -74.1320], "TUCURINCA": [10.5842, -74.1489],
            "PALOMAR": [10.7210, -74.1150], "LA CEIBA": [10.7350, -74.1620], "CAÑO MOCHO": [10.7820, -74.1850],
            "PALOMINO": [11.2442, -73.5623], "BURITACA": [11.2420, -73.7650], "GUACAMAYAL": [10.7292, -74.1594],
            "SEVILLA": [10.7667, -74.1500], "RIO FRIO": [10.9000, -74.1667], "FUNDACION": [10.5208, -74.1833]
        }

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🛰️ ENCENDER RADAR METEOROLÓGICO Y EPIDEMIOLÓGICO", type="primary", use_container_width=True):
            with st.spinner("Decodificando satélites e imprimiendo nombres en el terreno..."):
                
                df_t1 = cargar_historico_t1()
                if df_t1.empty:
                    st.error("🚨 No se pudo conectar a la base de datos operativa.")
                    st.stop()

                dict_poligonos_kml = {}
                if archivos_kml:
                    for f_kml in archivos_kml:
                        nombre_finca_kml = f_kml.name.upper().replace(".KML", "").strip()
                        poligonos = extraer_poligonos_kml(f_kml.read())
                        if poligonos:
                            dict_poligonos_kml[nombre_finca_kml] = poligonos

                fincas_unicas = df_t1['FINCA_NOM'].unique()
                analisis_fincas = []
                
                # 🎯 OPTIMIZACIÓN VIP: Pre-calcular el clima por coordenadas únicas para no saturar la API
                sectores_unicos = df_t1['SECTOR_NOM'].unique()
                cache_clima = {}
                for sec in sectores_unicos:
                    if sec in coor_estimadas:
                        gps = coor_estimadas[sec]
                        lluvia_90, lluvia_30 = consultar_clima_satelital(gps[0], gps[1])
                        cache_clima[sec] = (lluvia_90, lluvia_30)
                
                for finca in fincas_unicas:
                    if not finca or finca in ["NAN", "NONE", ""]: continue
                    
                    df_finca = df_t1[df_t1['FINCA_NOM'] == finca].sort_values(by='FECHA_DT')
                    fechas_vuelos = df_finca['FECHA_DT'].unique()
                    
                    sectores_frecuentes = df_finca['SECTOR_NOM'].value_counts()
                    sector_asociado = sectores_frecuentes.index[0] if not sectores_frecuentes.empty else "DESCONOCIDO"
                    
                    if len(fechas_vuelos) >= 2:
                        ultimo_vuelo = pd.to_datetime(fechas_vuelos[-1])
                        vuelo_anterior = pd.to_datetime(fechas_vuelos[-2])
                        dias_ciclo = (ultimo_vuelo - vuelo_anterior).days
                    else:
                        dias_ciclo = 30 
                    
                    if dias_ciclo <= 12:
                        estado = "🚨 CRÍTICO"
                        color_hex = "#cc0000"
                    elif dias_ciclo <= 20:
                        estado = "🟠 MODERADO"
                        color_hex = "#ff9900"
                    else:
                        estado = "🟢 CONTROLADO"
                        color_hex = "#27AE60"

                    gps = coor_estimadas.get(sector_asociado, [10.7483, -74.1542])
                    
                    # Usar la caché local optimizada
                    if sector_asociado in cache_clima:
                        lluvia_90d, lluvia_30d = cache_clima[sector_asociado]
                    else:
                        lluvia_90d, lluvia_30d = consultar_clima_satelital(gps[0], gps[1])
                    
                    alerta_epidemia = "Baja / Normal"
                    if lluvia_30d > 45.0: 
                        alerta_epidemia = "⚡ ALTA (Peligro Inminente)"

                    analisis_fincas.append({
                        "FINCA": finca,
                        "SECTOR": sector_asociado,
                        "ÚLTIMO RETORNO (Días)": float(dias_ciclo),
                        "ESTADO CICLO": estado,
                        "LLUVIA 90D (mm)": float(lluvia_90d),
                        "LLUVIA 30D (mm)": float(lluvia_30d),
                        "PRESIÓN HONGO": alerta_epidemia,
                        "COOR": gps,
                        "COLOR": color_hex
                    })

                # ==================================================
                # 💎 TARJETAS KPI SUPERIORES
                # ==================================================
                st.markdown("---")
                df_resumen = pd.DataFrame(analisis_fincas)
                
                if not df_resumen.empty:
                    total_fincas = len(df_resumen)
                    fincas_criticas = len(df_resumen[df_resumen['ESTADO CICLO'] == "🚨 CRÍTICO"])
                    fincas_hongo = len(df_resumen[df_resumen['PRESIÓN HONGO'].str.contains("ALTA")])

                    k1, k2, k3 = st.columns(3)
                    with k1: st.markdown(tarjeta_kpi("Total Fincas Mapeadas", f"{total_fincas} Fincas"), unsafe_allow_html=True)
                    with k2: st.markdown(tarjeta_kpi("Fincas con Ciclo Crítico (<12 Días)", f"{fincas_criticas} Fincas", "Atención Requerida", "#ff4b4b" if fincas_criticas > 0 else "#28a745"), unsafe_allow_html=True)
                    with k3: st.markdown(tarjeta_kpi("Fincas con Alta Presión de Hongo", f"{fincas_hongo} Fincas", "Por Lluvia", "#ff4b4b" if fincas_hongo > 0 else "#28a745"), unsafe_allow_html=True)

                # ==================================================
                # 🛰️ RENDERIZADO DEL MAPA SATELITAL
                # ==================================================
                mapa_magdalena = folium.Map(
                    location=[10.7483, -74.1542], 
                    zoom_start=10, 
                    tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
                    attr='Esri World Imagery'
                )
                st.markdown("### 🛰️ Mapa Georeferenciado en Vivo (Satelital)")
                
                sectores_dibujados = []
                kmls_usados = set() 

                for f_info in analisis_fincas:
                    finca_nom = f_info["FINCA"]
                    sector_nom = f_info["SECTOR"]
                    color_nodo = f_info["COLOR"]
                    
                    popup_text = f"""
                    <b>Finca:</b> {finca_nom} ({sector_nom})<br>
                    <b>Retorno:</b> {f_info["ÚLTIMO RETORNO (Días)"]:.0f} Días<br>
                    <b>Estado:</b> {f_info["ESTADO CICLO"]}<br>
                    <b>Lluvia Trimestre:</b> {f_info["LLUVIA 90D (mm)"]:.1f} mm<br>
                    <b>Lluvia Mensual:</b> {f_info["LLUVIA 30D (mm)"]:.1f} mm
                    """
                    
                    f_norm = limpiar_nombre(finca_nom)
                    kml_clave = None
                    
                    for k in dict_poligonos_kml.keys():
                        k_norm = limpiar_nombre(k)
                        if (k_norm in f_norm or f_norm in k_norm) and len(k_norm) > 3:
                            kml_clave = k
                            kmls_usados.add(k) 
                            break
                    
                    if kml_clave:
                        lats_finca = []
                        lons_finca = []
                        
                        for poligono in dict_poligonos_kml[kml_clave]:
                            folium.Polygon(
                                locations=poligono,
                                color=color_nodo,
                                weight=2,
                                fill=True,
                                fill_color=color_nodo,
                                fill_opacity=0.6,
                                tooltip=f"Finca: {finca_nom} | Estado: {f_info['ESTADO CICLO']}",
                                popup=folium.Popup(popup_text, max_width=300)
                            ).add_to(mapa_magdalena)
                            
                            lats_finca.extend([p[0] for p in poligono])
                            lons_finca.extend([p[1] for p in poligono])
                            
                        if lats_finca and lons_finca:
                            centro_lat = (min(lats_finca) + max(lats_finca)) / 2
                            centro_lon = (min(lons_finca) + max(lons_finca)) / 2
                            
                            html_label = f"""
                            <div style="
                                font-size: 11px; 
                                font-weight: 900; 
                                color: #FFFFFF; 
                                text-shadow: 2px 2px 3px #000, -2px -2px 3px #000, 2px -2px 3px #000, -2px 2px 3px #000, 0px 0px 5px #000;
                                white-space: nowrap;
                                text-align: center;
                                transform: translate(-50%, -50%);
                            ">
                                {finca_nom}
                            </div>
                            """
                            folium.Marker(
                                location=[centro_lat, centro_lon],
                                icon=folium.DivIcon(html=html_label)
                            ).add_to(mapa_magdalena)

                    else:
                        if sector_nom not in sectores_dibujados:
                            folium.CircleMarker(
                                location=f_info["COOR"],
                                radius=15,
                                color=color_nodo,
                                fill=True,
                                fill_color=color_nodo,
                                fill_opacity=0.8,
                                tooltip=f"Sector: {sector_nom}",
                                popup=folium.Popup(f"Sector: {sector_nom} (Suba KML para detalle)", max_width=300)
                            ).add_to(mapa_magdalena)
                            
                            html_label = f"""
                            <div style="font-size: 12px; font-weight: 900; color: #FFFFFF; text-shadow: 2px 2px 3px #000, -2px -2px 3px #000, 2px -2px 3px #000, -2px 2px 3px #000;">
                                {sector_nom}
                            </div>
                            """
                            folium.Marker(
                                location=[f_info["COOR"][0] + 0.01, f_info["COOR"][1]],
                                icon=folium.DivIcon(html=html_label)
                            ).add_to(mapa_magdalena)
                            
                            sectores_dibujados.append(sector_nom)

                # Dibuja KMLs huérfanos
                for kml_clave, poligonos in dict_poligonos_kml.items():
                    if kml_clave not in kmls_usados:
                        lats_finca = []
                        lons_finca = []
                        color_gris = "#A0A0A0" 
                        
                        for poligono in poligonos:
                            folium.Polygon(
                                locations=poligono,
                                color=color_gris,
                                weight=2,
                                fill=True,
                                fill_color=color_gris,
                                fill_opacity=0.4,
                                tooltip=f"Finca: {kml_clave} | Sin historial reciente",
                                popup=folium.Popup(f"<b>{kml_clave}</b><br>No se encontraron vuelos recientes en la base de datos.", max_width=300)
                            ).add_to(mapa_magdalena)
                            
                            lats_finca.extend([p[0] for p in poligono])
                            lons_finca.extend([p[1] for p in poligono])
                            
                        if lats_finca and lons_finca:
                            centro_lat = (min(lats_finca) + max(lats_finca)) / 2
                            centro_lon = (min(lons_finca) + max(lons_finca)) / 2
                            html_label = f"""
                            <div style="font-size: 10px; font-weight: 700; color: #CCCCCC; text-shadow: 1px 1px 2px #000; text-align: center; transform: translate(-50%, -50%);">
                                {kml_clave} (Inactiva)
                            </div>
                            """
                            folium.Marker(
                                location=[centro_lat, centro_lon],
                                icon=folium.DivIcon(html=html_label)
                            ).add_to(mapa_magdalena)

                # Renderiza el mapa en Streamlit
                st.components.v1.html(mapa_magdalena._repr_html_(), height=650)

                # ==================================================
                # 📋 TABLA DE REPORTE EPIDEMIOLÓGICO Y DESCARGA
                # ==================================================
                st.markdown("<br>### 📋 Reporte Epidemiológico y Satelital por Finca", unsafe_allow_html=True)
                
                df_vista = df_resumen.drop(columns=['COOR', 'COLOR']).copy()
                df_vista = df_vista.sort_values(by=['ESTADO CICLO', 'LLUVIA 30D (mm)'], ascending=[True, False])
                
                def pintar_estado_ciclo(row):
                    if "CRÍTICO" in row['ESTADO CICLO']: return ['background-color: #ffe6e6; color: #cc0000; font-weight:bold;'] * len(row)
                    if "MODERADO" in row['ESTADO CICLO']: return ['background-color: #fff3cd; color: #ff9900; font-weight:bold;'] * len(row)
                    return ['color: #155724;'] * len(row)

                # 💎 TABLA EJECUTIVA
                st.dataframe(
                    df_vista.style.apply(pintar_estado_ciclo, axis=1), 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "FINCA": st.column_config.TextColumn("📍 FINCA", width="medium"),
                        "SECTOR": st.column_config.TextColumn("🗺️ SECTOR", width="small"),
                        "ÚLTIMO RETORNO (Días)": st.column_config.NumberColumn("⏱️ RETORNO (DÍAS)", format="%.0f Días", width="small"),
                        "ESTADO CICLO": st.column_config.TextColumn("🚨 ESTADO", width="small"),
                        "LLUVIA 90D (mm)": st.column_config.NumberColumn("🌧️ LLUVIA 90D", format="%.1f mm", width="small"),
                        "LLUVIA 30D (mm)": st.column_config.NumberColumn("🌧️ LLUVIA 30D", format="%.1f mm", width="small"),
                        "PRESIÓN HONGO": st.column_config.TextColumn("🍄 RIESGO EPIDEMIOLÓGICO", width="medium")
                    }
                )
                
                excel_export = generar_excel_vip(df_vista, "Mapa_Agronomico")
                st.download_button(
                    label="💾 DESCARGAR REPORTE AGRONÓMICO (EXCEL VIP)", 
                    data=excel_export, 
                    file_name=f"Reporte_Agronomico_Satelital.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )

if __name__ == "__main__":
    pass
