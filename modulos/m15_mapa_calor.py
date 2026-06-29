import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import io
import json
import folium # Viene integrado o es sumamente ligero
import requests # Para conectarse al satélite del clima

# --- 🔌 CONEXIÓN Y REUTILIZACIÓN DE CACHÉ DE M14 ---
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

# 🧠 EXTRACTOR DE COORDENADAS KML CON ESCUDO DE NAMESPACES
def extraer_poligono_kml(kml_bytes):
    try:
        texto = kml_bytes.decode("utf-8")
        # Localiza la etiqueta de coordenadas usando expresiones regulares para evitar fallos de XML
        match = re.search(r'<coordinates>(.*?)</coordinates>', texto, re.DOTALL)
        if match:
            coor_str = match.group(1).strip()
            puntos = []
            for item in coor_str.split():
                partes = item.split(',')
                if len(partes) >= 2:
                    lon = float(partes[0])
                    lat = float(partes[1])
                    puntos.append([lat, lon]) # Folium exige [Lat, Lon]
            return puntos
    except: pass
    return None

# 🛰️ CONEXIÓN SATELITAL METEOROLÓGICA (Open-Meteo Satellites)
def consultar_lluvia_satelital(lat, lon):
    try:
        # Consulta el histórico de precipitaciones de los últimos 28 días
        hoy = datetime.now().date()
        hace_28_dias = hoy - timedelta(days=28)
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&start_date={hace_28_dias}&end_date={hoy}&daily=rain_sum&timezone=America/Bogota"
        res = requests.get(url, timeout=5).json()
        if "daily" in res and "rain_sum" in res["daily"]:
            lluvias = [float(x) for x in res["daily"]["rain_sum"] if x is not None]
            return sum(lluvias)
    except: pass
    return 0.0

@st.cache_data(show_spinner=False, ttl=3600)
def cargar_historico_t1():
    gc = inicializar_cliente_gspread()
    if not gc: return pd.DataFrame()
    boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
    t1_vals = boveda.worksheet("TABLA 1").get_all_values()
    df_t1 = pd.DataFrame(t1_vals[5:], columns=[str(c).upper().strip() for c in t1_vals[4]])
    
    col_fecha = next((c for c in df_t1.columns if 'FECHA' in c), 'FECHA')
    col_ha = next((c for c in df_t1.columns if 'NETA' in c or 'FUMIG' in c or 'HECT' in c), None)
    col_sector = next((c for c in df_t1.columns if 'SECTOR' in c), 'SECTOR')
    
    df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
    df_t1 = df_t1.dropna(subset=['FECHA_DT'])
    df_t1['HA_CALCULO'] = df_t1[col_ha].apply(a_numero_limpio)
    df_t1['SECTOR_NOM'] = df_t1[col_sector].astype(str).str.upper().str.strip()
    return df_t1

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-agronomo { color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }
    .card-meteo { background-color: #f8f9fa; border-left: 5px solid #2980b9; padding: 12px; border-radius: 5px; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-agronomo'>🗺️ Módulo 15: Radar Epidemiológico y Satelital</h1>", unsafe_allow_html=True)
    st.write("Análisis de ciclos de retorno por Sector cruzado con acumulados de lluvia satelital (Ventana de Alerta 21-29 días).")

    # --- 📥 CARGADOR TÁCTICO DE KML ---
    st.markdown("### 📂 1. Inyección de Polígonos de Precisión (Opcional)")
    archivos_kml = st.file_uploader("Arrastre aquí los archivos .kml de sus fincas o sectores para dibujar la topografía real", type=['kml'], accept_multiple_files=True)

    # Coordenadas maestras estimadas del Magdalena (Zona Bananera) por si no hay KML
    coor_estimadas = {
        "ORIHUECA": [10.7483, -74.1542], "FLORIDA": [10.7650, -74.1320], "TUCURINCA": [10.5842, -74.1489],
        "PALOMAR": [10.7210, -74.1150], "LA CEIBA": [10.7350, -74.1620], "CAÑO MOCHO": [10.7820, -74.1850],
        "PALOMINO": [11.2442, -73.5623], "BURITACA": [11.2420, -73.7650], "GUACAMAYAL": [10.7292, -74.1594]
    }

    if st.button("🛰️ ENCENDER RADAR METEOROLÓGICO Y EPIDEMIOLÓGICO", type="primary", use_container_width=True):
        with st.spinner("Analizando ciclos biológicos y conectando con satélites del clima..."):
            
            df_t1 = cargar_historico_t1()
            if df_t1.empty:
                st.error("No se pudo conectar a la base de datos de vuelos.")
                return

            # Procesar KMLs cargados
            dict_poligonos_kml = {}
            if archivos_kml:
                for f_kml in archivos_kml:
                    nombre_finca = f_kml.name.upper().replace(".KML", "").strip()
                    poligono = extraer_poligono_kml(f_kml.read())
                    if poligono:
                        dict_poligonos_kml[nombre_finca] = poligono

            # --- 🕵️‍♂️ CÁLCULO DE CICLOS DE RETORNO REALES ---
            sectores_unicos = df_t1['SECTOR_NOM'].unique()
            analisis_sectores = []

            for sector in sectores_unicos:
                if not sector or sector in ["NAN", "NONE", ""]: continue
                
                df_sec = df_t1[df_t1['SECTOR_NOM'] == sector].sort_values(by='FECHA_DT')
                fechas_vuelos = df_sec['FECHA_DT'].unique()
                
                # Calcular días del último ciclo cerrado
                if len(fechas_vuelos) >= 2:
                    ultimo_vuelo = pd.to_datetime(fechas_vuelos[-1])
                    vuelo_anterior = pd.to_datetime(fechas_vuelos[-2])
                    dias_ciclo = (ultimo_vuelo - vuelo_anterior).days
                else:
                    dias_ciclo = 30 # Por defecto si es nuevo
                
                # Regla de Oro del Comandante
                if dias_ciclo <= 12:
                    estado = "🚨 CRÍTICO (Presión Alta)"
                    color_hex = "#cc0000" # Rojo
                elif dias_ciclo <= 20:
                    estado = "🟠 MODERADO (Alerta)"
                    color_hex = "#ff9900" # Naranja
                else:
                    estado = "🟢 CONTROLADO (Óptimo)"
                    color_hex = "#27AE60" # Verde

                # Obtener coordenadas para el satélite
                gps = coor_estimadas.get(sector, [10.7483, -74.1542]) # Default Orihueca si es desconocido
                
                # Llamada al satélite meteorológico
                lluvia_acumulada = consultar_lluvia_satelital(gps[0], gps[1])
                
                # Ventana Predictiva del hongo (21-29 días después de lluvias)
                alerta_epidemia = "Baja"
                if lluvia_acumulada > 45.0: # Si llovieron más de 45mm en el mes
                    alerta_epidemia = "⚡ ALTA (Hongo en incubación: Ventana 21 días activa)"

                analisis_sectores.append({
                    "SECTOR": sector,
                    "ÚLTIMO RETORNO": f"{dias_ciclo} Días",
                    "ESTADO BIOLÓGICO": estado,
                    "LLUVIA 28 DÍAS (SATÉLITE)": f"{lluvia_acumulada:.1f} mm",
                    "ALERTA EPIDEMIOLÓGICA 21D": alerta_epidemia,
                    "COOR": gps,
                    "COLOR": color_hex
                })

            # --- 🗺️ CONSTRUCCIÓN DEL MAPA TÁCTICO ---
            # Centro del mapa en la Zona Bananera del Magdalena
            mapa_magdalena = folium.Map(location=[10.7483, -74.1542], zoom_start=10, tiles="OpenStreetMap")

            st.markdown("### 🛰️ Radar Visual en Vivo (Magdalena)")
            
            for s_info in analisis_sectores:
                nom_sec = s_info["SECTOR"]
                coor_node = s_info["COOR"]
                color_nodo = s_info["COLOR"]
                
                popup_text = f"""
                <b>Sector:</b> {nom_sec}<br>
                <b>Retorno:</b> {s_info["ÚLTIMO RETORNO"]}<br>
                <b>Estado:</b> {s_info["ESTADO BIOLÓGICO"]}<br>
                <b>Agua Satélite:</b> {s_info["LLUVIA 28 DÍAS (SATÉLITE)"]}<br>
                <b>Presión Sigatoka:</b> {s_info["ALERTA EPIDEMIOLÓGICA 21D"]}
                """
                
                # 💥 SI EL USUARIO SUBIÓ EL KML, DIBUJA EL POLÍGONO FINCA A FINCA 💥
                kml_clave = next((k for k in dict_poligonos_kml.keys() if k in nom_sec or nom_sec in k), None)
                if kml_clave:
                    folium.Polygon(
                        locations=dict_poligonos_kml[kml_clave],
                        color=color_nodo,
                        fill=True,
                        fill_color=color_nodo,
                        fill_opacity=0.4,
                        popup=folium.Popup(popup_text, max_width=300)
                    ).add_to(mapa_magdalena)
                else:
                    # Contingencia: Dibuja un domo de calor circular sobre el sector
                    folium.CircleMarker(
                        location=coor_node,
                        radius=20,
                        color=color_nodo,
                        fill=True,
                        fill_color=color_nodo,
                        fill_opacity=0.6,
                        popup=folium.Popup(popup_text, max_width=300)
                    ).add_to(mapa_magdalena)

            # Renderizado del mapa sin librerías externas usando el truco HTML
            st.components.v1.html(mapa_magdalena._repr_html_(), height=500)

            # --- 📋 TABLERO INFORMATIVO DE ALERTAS PREDICTIVAS ---
            st.markdown("<br>### 📋 Reporte de Alertas Tempranas (Efecto Lluvia + 21 Días)", unsafe_allow_html=True)
            df_resumen = pd.DataFrame(analisis_sectores).drop(columns=['COOR', 'COLOR'])
            
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
