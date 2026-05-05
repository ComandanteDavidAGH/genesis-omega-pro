import streamlit as st
import pandas as pd
from datetime import datetime

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 2. ARTILLERÍA VISUAL Y BLINDAJE CSS ---
arsenal_css = """
<style>
/* Ocultar marcas de agua de Streamlit */
[data-testid="stToolbarActions"] { display: none !important; }
.viewerBadge_container { display: none !important; }
footer { display: none !important; }

/* Colores y diseño corporativo */
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }

/* Títulos y Tarjetas */
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; font-size: 2.2rem; border-bottom: 3px solid #d4af37; padding-bottom: 10px; margin-bottom: 20px; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; border-left: 2px solid #e0e0e0; border-right: 2px solid #e0e0e0; border-bottom: 2px solid #e0e0e0; margin-bottom: 20px;}

/* Botones */
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; font-weight: bold !important; border-radius: 8px !important;}
button[kind="primary"]:hover { background-color: #d4af37 !important; color: #0d1b2a !important; border: 2px solid #0d1b2a !important; }

/* Subida de archivos */
[data-testid="stFileUploadDropzone"] { border: 2px dashed #0d1b2a !important; background-color: #ffffff !important; border-radius: 10px !important; }
</style>
"""
st.markdown(arsenal_css, unsafe_allow_html=True)

# --- 3. MENÚ TÁCTICO LATERAL ---
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #d4af37; font-family: Arial Black;'>🚀 GÉNESIS OMEGA</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; margin-top: -15px; color: white;'>Panel de Control AgroAéreo</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    menu = st.radio("🛰️ NAVEGACIÓN:", [
        "🏠 Centro de Mando", 
        "📥 1. Buzón de Carga (SAP & Pista)", 
        "⚙️ 2. Cruce y Validación Dosis", 
        "📊 3. Arqueo y Reportes", 
        "🛡️ Bóveda y Configuración"
    ])
    
    st.markdown("---")
    st.info(f"📅 Fecha Operativa:\n{datetime.now().strftime('%Y-%m-%d')}")

# --- 4. RUTAS DEL SISTEMA ---

if menu == "🏠 Centro de Mando":
    st.markdown("<h1 class='titulo-principal'>Centro de Mando | Operaciones</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3 style='color: #0d1b2a; margin-top:0;'>Bienvenido, Comandante.</h3>
        <p style='color: #333;'>El sistema <b>Génesis Omega Pro</b> está en línea y asegurado. Seleccione una operación en el menú lateral para iniciar el procesamiento de datos.</p>
        <ul>
            <li><b>Paso 1:</b> Cargue sus archivos de SAP y Pistas en el <i>Buzón de Carga</i>.</li>
            <li><b>Paso 2:</b> Ejecute el <i>Cruce y Validación</i> para asegurar dosis y precios.</li>
            <li><b>Paso 3:</b> Descargue sus <i>Arqueos y Reportes</i> listos para enviar.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

elif menu == "📥 1. Buzón de Carga (SAP & Pista)":
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje de Datos</h1>", unsafe_allow_html=True)
    st.write("Arrastre aquí la **Sábana de SAP** y los **Informes Diarios de Pista** para iniciar la sincronización.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("<h3 style='color:#0d1b2a;'>📁 1. Sábana Maestra SAP</h3>", unsafe_allow_html=True)
        archivo_sap = st.file_uploader("Suba el archivo EXPORT (Sábana de Inventario/Precios)", type=["xlsx", "xls", "csv"], key="sap_upload")
        if archivo_sap:
            st.success(f"✅ Sábana SAP lista: {archivo_sap.name}")

    with col2:
        st.markdown("<h3 style='color:#0d1b2a;'>🚁 2. Informes de Pista</h3>", unsafe_allow_html=True)
        archivos_pista = st.file_uploader("Suba los reportes diarios de las pistas", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pista_upload")
        if archivos_pista:
            st.success(f"✅ {len(archivos_pista)} reportes listos para consolidar.")
            
    st.markdown("---")
    col_btn, _ = st.columns([1, 2])
    with col_btn:
        if st.button("🚀 INICIAR LECTURA DE DATOS", type="primary", use_container_width=True):
            if archivo_sap and archivos_pista:
                with st.spinner("Decodificando información táctica..."):
                    try:
                        # 1. Masticar datos de SAP
                        if archivo_sap.name.endswith('.csv'):
                            df_sap = pd.read_csv(archivo_sap)
                        else:
                            df_sap = pd.read_excel(archivo_sap)
                        
                        st.session_state['df_sap'] = df_sap
                        
                        # 2. Consolidar Informes de Pista
                        lista_pistas = []
                        for pista in archivos_pista:
                            # Agregamos header=None para que no confunda el membrete con los títulos
                            if pista.name.endswith('.csv'):
                                df_temp = pd.read_csv(pista, header=None, encoding='utf-8', on_bad_lines='skip')
                            else:
                                df_temp = pd.read_excel(pista, header=None)
                            
                            # LIMPIEZA TÁCTICA: Borrar filas y columnas que estén 100% vacías
                            df_temp = df_temp.dropna(axis=1, how='all').dropna(axis=0, how='all')
                            
                            df_temp['ARCHIVO_ORIGEN'] = pista.name # Etiquetamos de qué pista viene
                            lista_pistas.append(df_temp)
                            
                        df_pistas_consol = pd.concat(lista_pistas, ignore_index=True)
                        
                        # Limpiamos los nombres de las columnas para que sean números genéricos temporalmente
                        df_pistas_consol.columns = [str(i) for i in range(len(df_pistas_consol.columns))]
                        
                        st.session_state['df_pistas'] = df_pistas_consol
                        
                        st.success("✅ ¡Datos devorados y pre-limpiados con éxito! Motores en línea.")
                        
                    except Exception as e:
                        st.error(f"🚨 Falla en los motores de lectura: {e}")
            else:
                st.error("🚨 Faltan suministros. Suba ambos frentes de datos.")

    # --- RADAR DE PREVISUALIZACIÓN ---
    if 'df_sap' in st.session_state and 'df_pistas' in st.session_state:
        st.markdown("### 👁️ Radar de Datos (Memoria RAM)")
        tab1, tab2 = st.tabs(["📁 Visión SAP", "🚁 Visión Pistas (Consolidado)"])
        
        with tab1:
            st.dataframe(st.session_state['df_sap'].head(10), use_container_width=True)
        with tab2:
            st.dataframe(st.session_state['df_pistas'].head(10), use_container_width=True)
            
elif menu == "⚙️ 2. Cruce y Validación Dosis":
    st.markdown("<h1 class='titulo-principal'>Validador Hiperespacial</h1>", unsafe_allow_html=True)
    st.warning("⚠️ Módulo en construcción. Aquí el sistema cruzará SAP con la Pista y aplicará las reglas de margen y precios de la matriz de Configuración.")

elif menu == "📊 3. Arqueo y Reportes":
    st.markdown("<h1 class='titulo-principal'>Central de Inteligencia</h1>", unsafe_allow_html=True)
    st.info("📊 Aquí se generarán los cuadros de diferencias, Excel de arqueos de fin de semana y reportes de sobrecostos dominicales listos para descargar.")

elif menu == "🛡️ Bóveda y Configuración":
    st.markdown("<h1 class='titulo-principal'>Bóveda Satelital</h1>", unsafe_allow_html=True)
    st.success("🔗 Estado de conexión con Google Sheets: Pendiente de enlace.")
    st.write("Desde aquí controlará los márgenes por tipo de productor, las marcas 'X' y los históricos de 3 años.")
