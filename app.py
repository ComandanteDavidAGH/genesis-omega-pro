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
                with st.spinner("Ejecutando Extracción Quirúrgica..."):
                    try:
                        # 1. Procesar SAP
                        if archivo_sap.name.endswith('.csv'):
                            df_sap = pd.read_csv(archivo_sap)
                        else:
                            df_sap = pd.read_excel(archivo_sap)
                        st.session_state['df_sap'] = df_sap
                        
                        # 2. Procesar Pistas con Escáner de "MEZCLA PREPARADA"
                        lista_mezclas = []
                        
                        for pista in archivos_pista:
                            if pista.name.endswith('.csv'):
                                df_raw = pd.read_csv(pista, header=None)
                            else:
                                df_raw = pd.read_excel(pista, header=None)
                            
                            # BUSCAR EL PUNTO DE INICIO: "MEZCLA PREPARADA"
                            mascara = df_raw.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                            
                            if mascara.any():
                                indice_inicio = mascara.idxmax()
                                # Tomamos desde ahí hasta el final
                                df_mezcla = df_raw.iloc[indice_inicio:].copy()
                                
                                # Limpieza de filas y columnas vacías
                                df_mezcla = df_mezcla.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                
                                # Etiquetamos el origen
                                df_mezcla['ARCHIVO_ORIGEN'] = pista.name
                                lista_mezclas.append(df_mezcla)
                        
                        if lista_mezclas:
                            df_pistas_consol = pd.concat(lista_mezclas, ignore_index=True)
                            # Normalizar nombres de columnas temporalmente
                            df_pistas_consol.columns = [str(i) for i in range(len(df_pistas_consol.columns))]
                            st.session_state['df_pistas'] = df_pistas_consol
                            st.success(f"✅ ¡Extracción exitosa! Se detectaron {len(lista_mezclas)} bloques de Mezcla Preparada.")
                        else:
                            st.error("🚨 No se encontró la frase 'MEZCLA PREPARADA' en los archivos de pista.")
                            
                    except Exception as e:
                        st.error(f"🚨 Error en la misión: {e}")
            else:
                st.error("🚨 Suministros incompletos. Cargue SAP y Pistas.")

    # --- RADAR DE PREVISUALIZACIÓN ---
    if 'df_sap' in st.session_state and 'df_pistas' in st.session_state:
        st.markdown("### 👁️ Radar de Datos (Memoria RAM)")
        
        # Agregamos contadores para su tranquilidad, Comandante
        filas_sap = len(st.session_state['df_sap'])
        filas_pistas = len(st.session_state['df_pistas'])
        
        col_c1, col_c2 = st.columns(2)
        col_c1.metric("Total Filas SAP", f"{filas_sap:,}")
        col_c2.metric("Total Filas Pistas", f"{filas_pistas:,}")

        tab1, tab2 = st.tabs(["📁 Visión SAP", "🚁 Visión Pistas (Consolidado)"])
        
        with tab1:
            st.write("Muestra de las primeras 50 filas (de la base completa):")
            st.dataframe(st.session_state['df_sap'].head(50), use_container_width=True)
        with tab2:
            st.write("Muestra de los bloques detectados:")
            st.dataframe(st.session_state['df_pistas'].head(50), use_container_width=True)

elif menu == "⚙️ 2. Cruce y Validación Dosis":
    st.markdown("<h1 class='titulo-principal'>Validador Hiperespacial</h1>", unsafe_allow_html=True)
    
    if 'df_sap' not in st.session_state or 'df_pistas' not in st.session_state:
        st.warning("⚠️ Radares apagados. Vaya al 'Buzón de Carga' y suba los suministros primero.")
    else:
        st.success("🟢 Suministros detectados en memoria. Motores de cruce listos.")
        
        if st.button("⚡ EJECUTAR EXTRACCIÓN DE DOSIS", type="primary", use_container_width=True):
            with st.spinner("Procesando matriz de productos..."):
                try:
                    df_raw = st.session_state['df_pistas']
                    datos_limpios = []
                    
                    # Buscamos la fila donde dice "PRODUCTO" para saber dónde empieza la lista
                    filas_producto = df_raw[df_raw.iloc[:, 1] == 'PRODUCTO'].index.tolist()
                    
                    for idx in filas_producto:
                        origen = df_raw.iloc[idx]['ARCHIVO_ORIGEN']
                        
                        # Recorremos desde la palabra "PRODUCTO" hacia abajo
                        fila_actual = idx + 1
                        while fila_actual < len(df_raw):
                            producto = str(df_raw.iloc[fila_actual, 1]).strip()
                            
                            # Si está vacío, terminamos este bloque
                            if producto.lower() == 'nan' or producto == '':
                                break
                                
                            cantidad = df_raw.iloc[fila_actual, 3] # Columna de cantidad
                            lote = df_raw.iloc[fila_actual, 4]     # Columna de Lote
                            
                            datos_limpios.append({
                                "PISTA_ORIGEN": origen,
                                "PRODUCTO": producto,
                                "CANTIDAD_PISTA": cantidad,
                                "LOTE_PISTA": lote
                            })
                            fila_actual += 1
                    
                    # Convertimos la lista limpia en un DataFrame
                    df_dosis_limpias = pd.DataFrame(datos_limpios)
                    st.session_state['df_dosis_limpias'] = df_dosis_limpias
                    
                    st.success("✅ ¡Extracción de Dosis completada!")
                    
                except Exception as e:
                    st.error(f"🚨 Falla en el escáner de dosis: {e}")

        # Mostrar el resultado de la limpieza
        if 'df_dosis_limpias' in st.session_state:
            st.markdown("### 📋 Tabla Oficial de Consumos (Lista para cruzar con SAP)")
            st.dataframe(st.session_state['df_dosis_limpias'], use_container_width=True)

elif menu == "📊 3. Arqueo y Reportes":
    st.markdown("<h1 class='titulo-principal'>Central de Inteligencia</h1>", unsafe_allow_html=True)
    st.info("📊 Aquí se generarán los cuadros de diferencias, Excel de arqueos de fin de semana y reportes de sobrecostos dominicales listos para descargar.")

elif menu == "🛡️ Bóveda y Configuración":
    st.markdown("<h1 class='titulo-principal'>Bóveda Satelital</h1>", unsafe_allow_html=True)
    st.success("🔗 Estado de conexión con Google Sheets: Pendiente de enlace.")
    st.write("Desde aquí controlará los márgenes por tipo de productor, las marcas 'X' y los históricos de 3 años.")
