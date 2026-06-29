import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import requests
import folium

# --- 🔌 CONEXIÓN Y UTILIDADES ---
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

def extraer_poligonos_kml(kml_bytes):
    try:
        texto = kml_bytes.decode("utf-8", errors="ignore")
        bloques = re.findall(r'<coordinates>(.*?)</coordinates>', texto, re.DOTALL)
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

# 🛰️ CONEXIÓN SATELITAL RECALIBRADA (Límite 90 Días para evadir bloqueo)
@st.cache_data(show_spinner=False, ttl=3600)
def consultar_clima_satelital(lat, lon):
    try:
        # Se solicitan past_days=90 (límite de la API gratuita) y usamos precipitation_sum
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&past_days=90&daily=precipitation_sum&timezone=America/Bogota"
        res = requests.get(url, timeout=10).json()
        
        if "daily" in res and "precipitation_sum" in res["daily"]:
            df_clima = pd.DataFrame({
                'fecha': pd.to_datetime(res['daily']['time']),
                'lluvia': [x if x is not None else 0.0 for x in res['daily']['precipitation_sum']]
            })
            
            hoy = pd.to_datetime(datetime.now().date())
            hace_30_dias = hoy - pd.Timedelta(days=30)
            
            # Filtramos para no contar el forecast futuro, solo el pasado
            lluvia_90d = df_clima[df_clima['fecha'] <= hoy]['lluvia'].sum()
            lluvia_30d = df_clima[(df_clima['fecha'] <= hoy) & (df_clima['fecha'] >= hace_30_dias)]['lluvia'].sum()
            
            return lluvia_90d, lluvia_30d
    except Exception as e: 
        print(f"Error clima: {e}")
        pass
    return 0.0, 0.0

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
    col_finca = next((c for c in df_t1.columns if 'FINCA' in c), 'FINCA')
    
    df_t1['FECHA_DT'] = df_t1[col_fecha].apply(procesar_fecha_pesada)
    df_t1 = df_t1.dropna(subset=['FECHA_DT'])
    df_t1['HA_CALCULO'] = df_t1[col_ha].apply(a_numero_limpio)
    df_t1['SECTOR_NOM'] = df_t1[col_sector].astype(str).str.upper().str.strip()
    df_t1['FINCA_NOM'] = df_t1[col_finca].astype(str).str.upper().str.strip()
    return df_t1

# --- 🚀 EJECUCIÓN PRINCIPAL ---
def ejecutar(purificar_lote, extraer_numero):
    st.markdown("""
    <style>
    .titulo-agronomo { color: #0d1b2a; border-bottom: 3px solid #27AE60; padding-bottom: 5px; font-family: 'Arial Black'; }
    div[data-testid="stDataFrame"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-agronomo'>🗺️ Módulo 15: Mapa de Calor Agronómico</h1>", unsafe_allow_html=True)
    st.write("Análisis de ciclos biológicos por FINCA y evaluación de lluvias satelitales.")

    st.markdown("### 📂 1. Inyección de Polígonos de Precisión")
    archivos_kml = st.file_uploader("Arrastre aquí los archivos .kml de sus fincas (Ej: ANGELES.kml)", type=['kml'], accept_multiple_files=True)

    coor_estimadas = {
        "ORIHUECA": [10.7483, -74.1542], "FLORIDA": [10.7650, -74.1320], "TUCURINCA": [10.5842, -74.1489],
        "PALOMAR": [10.7210, -74.1150], "LA CEIBA": [10.7350, -74.1620], "CAÑO MOCHO": [10.7820, -74.1850],
        "PALOMINO": [11.2442, -73.5623], "BURITACA": [11.2420, -73.7650], "GUACAMAYAL": [10.7292, -74.1594],
        "SEVILLA": [10.7667, -74.1500], "RIO FRIO": [10.9000, -74.1667]
    }

    if st.button("🛰️ ENCENDER RADAR METEOROLÓGICO Y EPIDEMIOLÓGICO", type="primary", use_container_width=True):
        with st.spinner("Decodificando satélites e imprimiendo nombres en el terreno..."):
            
            df_t1 = cargar_historico_t1()
            if df_t1.empty:
                st.error("No se pudo conectar a la base de datos operativa.")
                return

            dict_poligonos_kml = {}
            if archivos_kml:
                for f_kml in archivos_kml:
                    nombre_finca_kml = f_kml.name.upper().replace(".KML", "").strip()
                    poligonos = extraer_poligonos_kml(f_kml.read())
                    if poligonos:
                        dict_poligonos_kml[nombre_finca_kml] = poligonos

            fincas_unicas = df_t1['FINCA_NOM'].unique()
            analisis_fincas = []
            
            # --- 🕵️‍♂️ CÁLCULO DE CICLOS Y CLIMA POR FINCA ---
            for finca in fincas_unicas:
                if not finca or finca in ["NAN", "NONE", ""]: continue
                
                df_finca = df_t1[df_t1['FINCA_NOM'] == finca].sort_values(by='FECHA_DT')
                fechas_vuelos = df_finca['FECHA_DT'].unique()
                sector_asociado = df_finca['SECTOR_NOM'].iloc[-1]
                
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
                
                # OBTENER CLIMA REAL (90 Días y 30 Días)
                lluvia_90d, lluvia_30d = consultar_clima_satelital(gps[0], gps[1])
                
                alerta_epidemia = "BAJA"
                if lluvia_30d > 45.0: 
                    alerta_epidemia = "⚡ ALTA (Peligro Inminente)"

                analisis_fincas.append({
                    "FINCA": finca,
                    "SECTOR": sector_asociado,
                    "ÚLTIMO RETORNO": f"{dias_ciclo} Días",
                    "ESTADO": estado,
                    "LLUVIA 90D (mm)": lluvia_90d,
                    "LLUVIA 30D (mm)": lluvia_30d,
                    "PRESIÓN HONGO": alerta_epidemia,
                    "COOR": gps,
                    "COLOR": color_hex
                })

            # --- 🗺️ CONSTRUCCIÓN DEL MAPA TÁCTICO CON NOMBRES ---
            mapa_magdalena = folium.Map(location=[10.7483, -74.1542], zoom_start=10, tiles="OpenStreetMap")
            st.markdown("### 🛰️ Mapa Georeferenciado en Vivo")
            
            sectores_dibujados = []

            for f_info in analisis_fincas:
                finca_nom = f_info["FINCA"]
                sector_nom = f_info["SECTOR"]
                color_nodo = f_info["COLOR"]
                
                popup_text = f"""
                <b>Finca:</b> {finca_nom} ({sector_nom})<br>
                <b>Retorno:</b> {f_info["ÚLTIMO RETORNO"]}<br>
                <b>Estado:</b> {f_info["ESTADO"]}<br>
                <b>Lluvia Trimestre:</b> {f_info["LLUVIA 90D (mm)"]:.1f} mm<br>
                <b>Lluvia Mensual:</b> {f_info["LLUVIA 30D (mm)"]:.1f} mm
                """
                
                kml_clave = next((k for k in dict_poligonos_kml.keys() if k in finca_nom or finca_nom in k), None)
                
                if kml_clave:
                    for poligono in dict_poligonos_kml[kml_clave]:
                        # 1. Dibujar el polígono de la finca
                        folium.Polygon(
                            locations=poligono,
                            color=color_nodo,
                            weight=2,
                            fill=True,
                            fill_color=color_nodo,
                            fill_opacity=0.6,
                            tooltip=f"Finca: {finca_nom} | Estado: {f_info['ESTADO']}",
                            popup=folium.Popup(popup_text, max_width=300)
                        ).add_to(mapa_magdalena)
                        
                        # 2. 💥 CENTROIDE: Escribir el nombre de forma permanente en el mapa 💥
                        try:
                            lats = [p[0] for p in poligono]
                            lons = [p[1] for p in poligono]
                            centro_lat = sum(lats) / len(lats)
                            centro_lon = sum(lons) / len(lons)
                            
                            html_label = f"""
                            <div style="
                                font-size: 11px; 
                                font-weight: 900; 
                                color: black; 
                                text-shadow: 2px 2px 4px white, -2px -2px 4px white, 2px -2px 4px white, -2px 2px 4px white;
                                white-space: nowrap;
                            ">
                                {finca_nom}
                            </div>
                            """
                            folium.Marker(
                                location=[centro_lat, centro_lon],
                                icon=folium.DivIcon(html=html_label, icon_anchor=(20, 10))
                            ).add_to(mapa_magdalena)
                        except: pass

                else:
                    # Sin KML: Dibujar un punto y su etiqueta
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
                        <div style="font-size: 12px; font-weight: bold; color: black; text-shadow: 1px 1px 2px white;">
                            {sector_nom}
                        </div>
                        """
                        folium.Marker(
                            location=[f_info["COOR"][0] + 0.01, f_info["COOR"][1]],
                            icon=folium.DivIcon(html=html_label)
                        ).add_to(mapa_magdalena)
                        
                        sectores_dibujados.append(sector_nom)

            st.components.v1.html(mapa_magdalena._repr_html_(), height=600)

            # --- 📋 TABLERO DE ALERTAS ---
            st.markdown("<br>### 📋 Reporte Epidemiológico y Satelital por Finca", unsafe_allow_html=True)
            
            df_resumen = pd.DataFrame(analisis_fincas).drop(columns=['COOR', 'COLOR'])
            df_resumen = df_resumen.sort_values(by=['ESTADO', 'LLUVIA 30D (mm)'], ascending=[True, False])
            
            # Formato de la tabla (Cambiado de Lluvia Año a Lluvia 90D)
            df_resumen['LLUVIA 90D (mm)'] = df_resumen['LLUVIA 90D (mm)'].apply(lambda x: f"{x:.1f} mm")
            df_resumen['LLUVIA 30D (mm)'] = df_resumen['LLUVIA 30D (mm)'].apply(lambda x: f"{x:.1f} mm")

            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
