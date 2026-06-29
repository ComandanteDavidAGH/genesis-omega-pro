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

# 🛰️ CONEXIÓN SATELITAL CLIMÁTICA
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
    except Exception as e: pass
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
    st.write("Análisis de ciclos biológicos por FINCA sobre terreno satelital y lluvia trimestral.")

    # --- 💥 GESTOR DE CARGA BLINDADO (Sin columnas) 💥 ---
    st.markdown("### 📂 1. Inyección de Polígonos de Precisión")
    
    if "kml_reset_key" not in st.session_state:
        st.session_state.kml_reset_key = 0

    archivos_kml = st.file_uploader(
        "Arrastre aquí los archivos .kml de sus fincas", 
        type=['kml'], 
        accept_multiple_files=True, 
        key=f"kml_uploader_{st.session_state.kml_reset_key}"
    )

    if st.button("🗑️ Vaciar Bandeja de KMLs", type="secondary"):
        st.session_state.kml_reset_key += 1
        st.rerun()

    coor_estimadas = {
        "ORIHUECA": [10.7483, -74.1542], "FLORIDA": [10.7650, -74.1320], "TUCURINCA": [10.5842, -74.1489],
        "PALOMAR": [10.7210, -74.1150], "LA CEIBA": [10.7350, -74.1620], "CAÑO MOCHO": [10.7820, -74.1850],
        "PALOMINO": [11.2442, -73.5623], "BURITACA": [11.2420, -73.7650], "GUACAMAYAL": [10.7292, -74.1594],
        "SEVILLA": [10.7667, -74.1500], "RIO FRIO": [10.9000, -74.1667], "FUNDACION": [10.5208, -74.1833]
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
                <b>Retorno:</b> {f_info["ÚLTIMO RETORNO"]}<br>
                <b>Estado:</b> {f_info["ESTADO"]}<br>
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
                            tooltip=f"Finca: {finca_nom} | Estado: {f_info['ESTADO']}",
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

            st.components.v1.html(mapa_magdalena._repr_html_(), height=650)

            st.markdown("<br>### 📋 Reporte Epidemiológico y Satelital por Finca", unsafe_allow_html=True)
            
            df_resumen = pd.DataFrame(analisis_fincas).drop(columns=['COOR', 'COLOR'])
            df_resumen = df_resumen.sort_values(by=['ESTADO', 'LLUVIA 30D (mm)'], ascending=[True, False])
            
            df_resumen['LLUVIA 90D (mm)'] = df_resumen['LLUVIA 90D (mm)'].apply(lambda x: f"{x:.1f} mm")
            df_resumen['LLUVIA 30D (mm)'] = df_resumen['LLUVIA 30D (mm)'].apply(lambda x: f"{x:.1f} mm")

            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
