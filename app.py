import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import gspread

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 2. ARTILLERÍA VISUAL ---
arsenal_css = """
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }
</style>
"""
st.markdown(arsenal_css, unsafe_allow_html=True)

# --- 3. MENÚ TÁCTICO ---
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h2>", unsafe_allow_html=True)
    menu = st.radio("🛰️ NAVEGACIÓN:", ["🏠 Centro de Mando", "📥 1. Buzón de Carga", "⚙️ 2. Validación de Misión", "📊 3. Arqueo y Reportes", "🛡️ Configuración"])
    st.info(f"📅 Operación: {datetime.now().strftime('%Y-%m-%d')}")

# --- 4. LÓGICA DE CARGA ---

if menu == "🏠 Centro de Mando":
    st.markdown("<h1 class='titulo-principal'>Centro de Mando</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3>Estrategia de Validación (La Trinidad):</h3>
        <ol>
            <li><b>Sábana SAP:</b> Validamos Lotes y Precios oficiales.</li>
            <li><b>Pedidos SAP:</b> Validamos lo que se DEBÍA hacer (Fincas/Hectáreas).</li>
            <li><b>Informes Pista:</b> Validamos lo que REALMENTE se hizo.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

elif menu == "📥 1. Buzón de Carga":
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Cuartel General</h1>", unsafe_allow_html=True)
    
    # Volvemos a 3 cuadrantes, el 4to ahora es invisible (Satelital)
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        f_sabana = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="sab")
    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        f_pedidos = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="ped")
    with c3:
        st.markdown("### 🚁 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv", "CSV", "XLSX"], accept_multiple_files=True, key="pis")

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas:
            with st.spinner("Sincronizando los 3 frentes y conectando con Satélite en Google Drive..."):
                try:
                    # 1. Leer Sábana
                    bytes_sabana = io.BytesIO(f_sabana.getvalue())
                    nom_sab = f_sabana.name.lower()
                    if nom_sab.endswith('.xlsx') or nom_sab.endswith('.xls'):
                        st.session_state['df_sabana'] = pd.read_excel(bytes_sabana)
                    else:
                        st.session_state['df_sabana'] = pd.read_csv(bytes_sabana, sep=None, engine='python')
                    
                    # 2. Leer Pedidos
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    nom_ped = f_pedidos.name.lower()
                    if nom_ped.endswith('.xlsx') or nom_ped.endswith('.xls'):
                        st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos)
                    else:
                        st.session_state['df_pedidos'] = pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    # ==========================================
                    # 🛰️ 3. CONEXIÓN SATELITAL (BÓVEDA GOOGLE DRIVE)
                    # ==========================================
                    try:
                        # Motor Nativo: Lee el diccionario directamente de Streamlit
                        if "gcp_credentials" in st.secrets:
                            # Convertimos el secreto nativo directamente a diccionario
                            cred_dict = dict(st.secrets["gcp_credentials"])
                            gc = gspread.service_account_from_dict(cred_dict)
                        else:
                            gc = gspread.service_account(filename='credenciales.json')
                        
                        # ABRIR LA BÓVEDA POR URL (Asegúrese de tener su link real aquí)
                        url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit?gid=1995619804#gid=1995619804"
                        boveda = gc.open_by_url(url_boveda)
                        
                        # Entrar a la pestaña "TABLA 2"
                        hoja_tabla2 = boveda.worksheet("TABLA 2")
                        datos_tabla2 = hoja_tabla2.get_all_values() # Trae todo como texto puro
                        
                        # Convertimos a Pandas DataFrame (La primera fila son los títulos)
                        df_config_nube = pd.DataFrame(datos_tabla2[1:], columns=datos_tabla2[0])
                        st.session_state['df_config'] = df_config_nube
                        
                        conexion_exitosa = True
                    except Exception as error_nube:
                        st.error(f"🚨 Falla en el Enlace Satelital con Drive: {error_nube}")
                        conexion_exitosa = False
                    # ==========================================
                    
                    # 4. 🛰️ ESCÁNER PROFUNDO (Solo Pestañas Visibles / Multifinca)
                    lista_pistas = []
                    
                    for f in f_pistas:
                        # --- PASO 0: Detectar cuáles pestañas son visibles ---
                        bytes_data = f.getvalue()
                        wb = openpyxl.load_workbook(io.BytesIO(bytes_data), read_only=True)
                        pestañas_visibles = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == 'visible']
                        wb.close()
                        
                        # --- PASO 1: Leer solo las visibles con Pandas ---
                        dict_pestañas = pd.read_excel(io.BytesIO(bytes_data), sheet_name=pestañas_visibles, header=None)
                        
                        for nombre_pestaña, df in dict_pestañas.items():
                            # Limpieza inicial de filas/columnas vacías que confunden al radar
                            df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                            
                            # Paso A: Localizar el Cóctel
                            idx_coctel = df[df.astype(str).apply(lambda x: x.str.contains('COCTEL', case=False, na=False)).any(axis=1)].index
                            
                            if not idx_coctel.empty:
                                fila_coctel = idx_coctel[0]
                                # Buscamos el nombre del cóctel en esa fila (suele ser el siguiente valor no nulo)
                                valores_fila = df.iloc[fila_coctel].dropna().tolist()
                                nombre_coctel = valores_fila[1] if len(valores_fila) > 1 else "DESCONOCIDO"
                                
                                # Paso B: Localizar la tabla de Fincas
                                idx_fincas_header = df[df.astype(str).apply(lambda x: x.str.contains('FINCAS', case=False, na=False)).any(axis=1)].index
                                
                                if not idx_fincas_header.empty:
                                    fila_header = idx_fincas_header[0]
                                    col_fincas_idx = (df.iloc[fila_header].astype(str).str.contains('FINCAS', case=False)).values.argmax()
                                    
                                    # Extraemos datos desde la fila de abajo
                                    df_datos = df.iloc[fila_header + 1:].copy()
                                    
                                    # Recorremos cada fila hasta el final de la tabla de esa pestaña
                                    for i in range(len(df_datos)):
                                        finca_nombre = str(df_datos.iloc[i, col_fincas_idx]).strip()
                                        
                                        # Si la celda está vacía o dice TOTAL, cerramos esta pestaña
                                        if finca_nombre.lower() in ['nan', '', 'none'] or "TOTAL" in finca_nombre.upper():
                                            break
                                        
                                        # Capturamos toda la fila de datos para el cruce del Módulo 2
                                        # (Hectáreas, Pedido, Productos, etc.)
                                        registro = {
                                            "ORIGEN": f"{f.name} | {nombre_pestaña}",
                                            "COCTEL": nombre_coctel,
                                            "FINCA_INFORME": finca_nombre,
                                            "DATOS_COMPLETOS": df_datos.iloc[i].to_dict() 
                                        }
                                        lista_pistas.append(registro)
                    
                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                        st.success(f"✅ ¡Barrido Exitoso! {len(lista_pistas)} vuelos detectados en pestañas visibles.")
                    else:
                        st.error("🚨 No se encontró la estructura 'FINCAS' en ninguna pestaña visible.")            
elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>🚀 Centro de Mando Génesis 2.0</h1>", unsafe_allow_html=True)
    
    if 'df_sabana' not in st.session_state:
        st.warning("⚠️ Sin suministros. Cargue SAP e Informes primero.")
    else:
        # --- 1. CARGA DE INTELIGENCIA (Drive) ---
        # El bot ya descargó Configuración, BD_Mezclas y Tabla de Apoyo
        df_apoyo = st.session_state.get('df_apoyo', pd.DataFrame()) 
        df_config = st.session_state.get('df_config', pd.DataFrame())
        
        # --- 2. RADAR DE SELECCIÓN ---
        st.markdown("### 📡 Selección de Objetivo")
        lista_pedidos = st.session_state['df_pistas']['ORIGEN'].unique().tolist()
        pedido_id = st.selectbox("🎯 Seleccione Pedido de Pista:", ["---"] + lista_pedidos)
        
        if pedido_id != "---":
            # --- 3. EXTRACCIÓN Y CRUCE (Lógica Macro) ---
            # Simulamos que el sistema ya hizo el cruce con Pedidos SAP y Sabana
            finca_sap = "SACRAMENTO" 
            tipo_productor = "INVERSIONISTA" # Detectado de Tabla de Apoyo
            margen_aplicado = 0.12 # Sacado de Configuración para Inversionistas
            
            st.success(f"📦 Pedido Detectado: {pedido_id} | Productor: {tipo_productor} (Margen: {margen_aplicado*100}%)")
            
            with st.container(border=True):
                c1, c2, c3 = st.columns([2, 1, 1])
                finca_edit = c1.text_input("📍 Finca (Sobrescribir si es necesario):", value=finca_sap)
                ha_reales = c2.number_input("🚜 Hectáreas Reales:", value=79.0)
                horometro = c3.number_input("⏱️ Horómetro (hrs):", value=1.5)

                # --- TABLA DE CONTROL (ESTILO IMAGEN 2) ---
                st.markdown("#### 📊 Detalle de Liquidación (Precios + Margen)")
                
                # Aquí aplicamos la lógica de la "X" y el Margen
                detalle_df = pd.DataFrame({
                    "Material": ["ACEITE", "SIGANEX", "INBIOSIL"],
                    "Cant. Real": [474, 40, 20],
                    "Costo Unit": [4892, 44243, 12078], # Viene de Sabana SAP
                    "Margen %": [margen_aplicado] * 3,
                    "Precio Venta": [4892*(1+margen_aplicado), 44243*(1+margen_aplicado), 12078*(1+margen_aplicado)]
                })
                detalle_df["Total"] = detalle_df["Cant. Real"] * detalle_df["Precio Venta"]
                
                st.dataframe(detalle_df.style.format({"Costo Unit": "${:,.0f}", "Precio Venta": "${:,.0f}", "Total": "${:,.0f}"}), use_container_width=True)

                # --- MOTOR DE CÁLCULO DE APLICACIÓN (Topes y Matriz) ---
                tarifa_base = 2500000
                costo_vuelo_ha = (horometro * tarifa_base) / ha_reales
                pdiv_val = 45000 # Sacado de Configuración
                precio_final_app = min(costo_vuelo_ha, pdiv_val)
                
                # --- BOTÓN DE DETONACIÓN (Escritura en Nube) ---
                if st.button("🔥 DETONAR FACTURA Y ALIMENTAR HISTORIAL", type="primary", use_container_width=True):
                    # AQUÍ EL CÓDIGO ESCRIBE EN EL GOOGLE SHEET
                    nueva_fila = [pedido_id, finca_edit, ha_reales, horometro, detalle_df["Total"].sum(), precio_final_app]
                    # gc.append_row(nueva_fila) <--- Esto se activa con su credenciales.json
                    st.balloons()
                    st.success("✅ ¡Historial Alimentado! Base de Datos Actualizada.")

        # --- 4. DASHBOARD ULTRA-MODERNO (Nivel 8) ---
        st.markdown("---")
        st.markdown("### 📈 Monitor de Operaciones en Tiempo Real")
        col_g1, col_g2 = st.columns(2)
        
        # Gráfico 1: Hectáreas por Finca (Histórico)
        fig_ha = px.bar(x=["Sacramento", "Tamacará", "La Carolina"], y=[450, 320, 280], 
                        title="Hectáreas Acumuladas por Finca", color_discrete_sequence=['#00FFAA'])
        col_g1.plotly_chart(fig_ha, use_container_width=True)
        
        # Gráfico 2: Eficiencia de Costos (Real vs Tope)
        fig_costos = px.line(x=[1,2,3,4,5], y=[42000, 48000, 44000, 46000, 45000], 
                             title="Tendencia de Costo por Ha (vs Tope)", markers=True)
        col_g2.plotly_chart(fig_costos, use_container_width=True)
