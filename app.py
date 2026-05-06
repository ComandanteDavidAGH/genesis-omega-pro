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
                    
                    # 4. Leer Pistas
                    lista_pistas = []
                    errores_pistas = []
                    
                    for f in f_pistas:
                        bytes_pista = io.BytesIO(f.getvalue())
                        nom_pis = f.name.lower()
                        
                        try:
                            if nom_pis.endswith('.csv'):
                                df_raw = pd.read_csv(bytes_pista, header=None, sep=None, engine='python', on_bad_lines='skip')
                                m = df_raw.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                if m.any():
                                    df_m = df_raw.iloc[m.idxmax():].copy()
                                    df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                    df_m['ORIGEN'] = f.name
                                    lista_pistas.append(df_m)
                                    
                            elif nom_pis.endswith('.xlsx'):
                                wb = openpyxl.load_workbook(bytes_pista, read_only=True, data_only=True)
                                visibles = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
                                
                                if visibles:
                                    bytes_pandas = io.BytesIO(f.getvalue())
                                    dict_p = pd.read_excel(bytes_pandas, sheet_name=visibles, header=None)
                                    for name, df in dict_p.items():
                                        m = df.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                        if m.any():
                                            df_m = df.iloc[m.idxmax():].copy()
                                            df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                            df_m['ORIGEN'] = f"{f.name} ({name})"
                                            lista_pistas.append(df_m)
                                            
                            else:
                                dict_p = pd.read_excel(bytes_pista, sheet_name=None, header=None)
                                for name, df in dict_p.items():
                                    m = df.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                    if m.any():
                                        df_m = df.iloc[m.idxmax():].copy()
                                        df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                        df_m['ORIGEN'] = f"{f.name} ({name})"
                                        lista_pistas.append(df_m)
                                        
                        except Exception as e_pista:
                            errores_pistas.append(f"{f.name} ({str(e_pista)})")
                    
                    if lista_pistas and conexion_exitosa:
                        st.session_state['df_pistas'] = pd.concat(lista_pistas, ignore_index=True)
                        st.success(f"✅ ¡Operación Exitosa! SAP: {len(st.session_state['df_sabana'])} filas | Pedidos: {len(st.session_state['df_pedidos'])} filas | Pistas: {len(lista_pistas)} bloques | 📡 Satélite TABLA 2: Conectado ({len(st.session_state['df_config'])} registros).")
                        
                        if errores_pistas:
                            st.warning(f"⚠️ Algunos archivos de pista fueron saltados por formato ilegible: {', '.join(errores_pistas)}")
                    elif not lista_pistas:
                        st.error("🚨 No se encontró información válida de 'MEZCLA PREPARADA' en las pistas.")
                        
                except Exception as e:
                    st.error(f"🚨 Error crítico en el ensamblaje principal: {e}")
        else:
            st.error("🚨 Faltan suministros locales. Suba los 3 frentes requeridos.")
            
elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>Centro de Mando Integral (Cruce Tripartito)</h1>", unsafe_allow_html=True)
    
    # 1. VERIFICACIÓN DE SUMINISTROS
    if 'df_sabana' not in st.session_state or 'df_pistas' not in st.session_state:
        st.warning("⚠️ Faltan suministros. Sincronice la Trinidad en el 'Buzón de Carga' primero.")
    else:
        df_pistas = st.session_state['df_pistas']
        
        # --- ESCUADRÓN ALFA: RADAR DE MISIONES ---
        st.markdown("### 📡 Radar de Vuelos Detectados")
        
        # Simulamos la extracción de Pedidos del Informe de Pista (Última columna del CSV)
        # Aquí recolectamos todos los vuelos únicos detectados
        lista_pedidos_detectados = ["Seleccione un Pedido..."]
        if 'ORIGEN' in df_pistas.columns:
            lista_pedidos_detectados.extend(df_pistas['ORIGEN'].unique().tolist())
        else:
            lista_pedidos_detectados.extend(["170035970 - SACRAMENTO 1 (Demo)", "170035971 - TAMACARA (Demo)"]) # Respaldo si no detecta la columna
            
        pedido_seleccionado = st.selectbox("🎯 Fije el blanco (Seleccione el Pedido a Facturar):", lista_pedidos_detectados)
        
        if pedido_seleccionado != "Seleccione un Pedido...":
            st.success(f"🟢 Blanco fijado: {pedido_seleccionado}. Desplegando telemetría...")
            
            # --- ESCUADRÓN BRAVO: DATOS TÁCTICOS Y LABORATORIO ---
            with st.form("form_laboratorio"):
                st.markdown("#### 1️⃣ Coordenadas de Vuelo (Modificables)")
                c1, c2, c3, c4 = st.columns(4)
                
                # Libertad táctica: Sobrescribir Finca
                finca_sugerida = "SACRAMENTO" # Esto vendrá del BUSCARV satelital
                finca_final = c1.text_input("Finca (Editable vs SAP):", value=finca_sugerida)
                
                hectareas_reales = c2.number_input("Hectáreas Reales:", value=79.0, step=0.1)
                coctel = c3.text_input("Cóctel (BD_MEZCLAS):", value="SGMN63+FE", disabled=True)
                dias_ciclo = c4.number_input("Días Ciclo (Auto):", value=10) # Se autocompleta con Tabla de Apoyo
                
                st.markdown("---")
                st.markdown("#### 2️⃣ Laboratorio de Dosis (Piloto vs SAP vs Teórica)")
                
                # Aquí se programa la lógica de sus Macros (Inbiosil, Acondicionador, "X")
                st.info("💡 **Reglas Activas:** Aplicando multiplicador 'X' | Evaluando Acondicionador vs Fertilizantes | Validando Inbiosil.")
                
                # Simulador visual de su cuadro de cruce
                datos_cruce = pd.DataFrame({
                    "PRODUCTO": ["ACEITE", "ADHERENTE", "ACONDICIONADOR", "SIGANEX", "INBIOSIL"],
                    "CANT. PILOTO": [474, 10.3, 1.6, 40, 20],
                    "PENDIENTE SAP": [474, 10.3, 1.6, 40, 20],
                    "DOSIS TEÓRICA (BD)": ["6.0 x ha", "0.13 x ha", "0.02 x ha", "0.5 x ha", "Varía por rol"],
                    "ESTADO MACRO": ["✅ Exacto", "✅ Exacto", "⚠️ Cambio x Fertilizante", "✅ Exacto", "🟢 Rol: Activo (X aplicada)"]
                })
                st.dataframe(datos_cruce, use_container_width=True)
                
                # --- ESCUADRÓN CHARLIE: MOTOR DE LIQUIDACIÓN ---
                st.markdown("---")
                st.markdown("#### 3️⃣ Motor de Liquidación (Horómetro y Tarifas)")
                
                col_maq1, col_maq2, col_maq3 = st.columns(3)
                tipo_maquina = col_maq1.radio("Vehículo de Asalto:", ["✈️ Avión", "🚁 Dron"])
                
                # HORÓMETRO MANUAL (Un solo campo, como usted ordenó)
                horometro = col_maq2.number_input("Horómetro Reportado (hrs):", value=1.5, step=0.1, min_value=0.0) if tipo_maquina == "✈️ Avión" else 0.0
                
                # MATRIZ DE MÚLTIPLES AVIONES
                multi_avion = col_maq3.checkbox("Aplicar Matriz (Múltiples Aviones)")
                porcentaje_prorrateo = col_maq3.number_input("Participación del Avión (%)", value=100.0, step=1.0) / 100.0 if multi_avion else 1.0
                
                btn_facturar = st.form_submit_button("🔥 CALCULAR Y FACTURAR", type="primary", use_container_width=True)
                
            # --- DETONACIÓN MATEMÁTICA ---
            if btn_facturar:
                with st.spinner("Procesando ecuaciones, macros y topes..."):
                    # Simulamos la lectura de la TABLA DE APOYO y Configuración
                    valor_hora_avion = 2500000
                    tarifa_dron_ha = 60000
                    tope_pista = 45000
                    es_pdiv = True if "PDIV" in finca_final.upper() else False
                    
                    if tipo_maquina == "✈️ Avión":
                        # CÁLCULO DE AVIÓN CON MATRIZ ((D19*B19)/A19)
                        costo_total_vuelo = (horometro * valor_hora_avion) * porcentaje_prorrateo
                        costo_por_ha = costo_total_vuelo / hectareas_reales if hectareas_reales > 0 else 0
                        
                        if es_pdiv:
                            precio_aplicacion = costo_por_ha
                            alerta_tope = "🟢 Regla PDIV: Tope Ignorado."
                        else:
                            precio_aplicacion = min(costo_por_ha, tope_pista)
                            alerta_tope = "🔴 Tope de Pista Aplicado." if costo_por_ha > tope_pista else "✅ Precio Real dentro del Tope."
                            
                        metrica_tiempo = f"{horometro} hrs"
                        
                    else:
                        # CÁLCULO DE DRON (Directo por Hectárea)
                        precio_aplicacion = tarifa_dron_ha
                        alerta_tope = "🚁 Tarifa Dron Fija."
                        metrica_tiempo = "N/A (Dron)"
                        costo_total_vuelo = precio_aplicacion * hectareas_reales

                    # RESULTADOS TÁCTICOS
                    st.markdown("### 🏆 RECIBO DE COMBATE")
                    st.info(f"**Finca Oficial:** {finca_final} | **Reglas:** {alerta_tope}")
                    
                    r1, r2, r3, r4 = st.columns(4)
                    r1.metric("🚜 Hectáreas Facturadas", f"{hectareas_reales:.1f} ha")
                    r2.metric("⏱️ Horómetro / Tiempo", metrica_tiempo)
                    r3.metric("💲 Tarifa Final (por ha)", f"${precio_aplicacion:,.0f}")
                    r4.metric("💰 TOTAL VUELO", f"${(precio_aplicacion * hectareas_reales):,.0f}")
