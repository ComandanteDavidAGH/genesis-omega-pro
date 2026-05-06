import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io

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
    
    c1, c2 = st.columns(2)
    c3, c4 = st.columns(2)
    
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        f_sabana = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="sab")
    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        f_pedidos = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="ped")
    with c3:
        st.markdown("### 🚁 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv", "CSV", "XLSX"], accept_multiple_files=True, key="pis")
    with c4:
        st.markdown("### ⚙️ 4. Base TABLA 2 / Config")
        f_config = st.file_uploader("Fincas, Topes y Productores", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="conf")

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas and f_config:
            with st.spinner("Sincronizando los 4 frentes de batalla..."):
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
                        
                    # 3. LECTURA DE TABLA 2 / CONFIGURACIÓN (Reubicado al lugar correcto)
                    bytes_config = io.BytesIO(f_config.getvalue())
                    nom_conf = f_config.name.lower()
                    if nom_conf.endswith('.xlsx') or nom_conf.endswith('.xls'):
                        st.session_state['df_config'] = pd.read_excel(bytes_config)
                    else:
                        st.session_state['df_config'] = pd.read_csv(bytes_config, sep=None, engine='python')
                    
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
                    
                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.concat(lista_pistas, ignore_index=True)
                        st.success(f"✅ ¡Cuartel Sincronizado! SAP: {len(st.session_state['df_sabana'])} | Pedidos: {len(st.session_state['df_pedidos'])} | Pistas: {len(lista_pistas)} | Config: {len(st.session_state['df_config'])} filas.")
                        
                        if errores_pistas:
                            st.warning(f"⚠️ Nota: Algunos archivos fueron saltados por formato ilegible: {', '.join(errores_pistas)}")
                    else:
                        st.error("🚨 No se encontró información válida de 'MEZCLA PREPARADA' en las pistas.")
                        
                except Exception as e:
                    st.error(f"🚨 Error crítico en el ensamblaje principal: {e}")
        else:
            st.error("🚨 Faltan suministros. Suba los 4 frentes.")
            
elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>Validación Cruzada (La Trinidad)</h1>", unsafe_allow_html=True)
    
    if 'df_sabana' not in st.session_state or 'df_pedidos' not in st.session_state or 'df_pistas' not in st.session_state:
        st.warning("⚠️ Faltan suministros. Vaya al 'Buzón de Carga' y sincronice la Trinidad primero.")
    else:
        st.success("🟢 Radares enlazados. Motores de validación listos.")
        
        if st.button("⚡ EJECUTAR CRUCE TÁCTICO", type="primary", use_container_width=True):
            with st.spinner("Cruzando coordenadas: Pistas vs Sábana vs Pedidos..."):
                try:
                    df_pistas = st.session_state['df_pistas']
                    df_sabana = st.session_state['df_sabana']
                    df_pedidos = st.session_state['df_pedidos']
                    
                    # 1. Preparar Sábana (Identificar columnas clave)
                    cols_sabana = [str(c).upper().strip() for c in df_sabana.columns]
                    df_sabana.columns = cols_sabana
                    col_prod_sab = next((c for c in cols_sabana if 'MATERIAL' in c or 'DESCRIPCI' in c), None)
                    col_lote_sab = next((c for c in cols_sabana if 'LOTE' in c), None)
                    
                    datos_validados = []
                    
                    # 2. Escáner de Pistas
                    # Buscamos dónde dice "PRODUCTO" en cualquier columna para saber dónde inicia la tabla
                    filas_producto = df_pistas[df_pistas.astype(str).apply(lambda x: x.str.contains('PRODUCTO', case=False, na=False)).any(axis=1)].index.tolist()
                    
                    for idx in filas_producto:
                        origen = df_pistas.iloc[idx]['ORIGEN'] if 'ORIGEN' in df_pistas.columns else "Desconocido"
                        
                        # Extraer Finca y Hectáreas buscando en las filas justo arriba de "PRODUCTO"
                        finca = "No detectada"
                        hectareas = "No detectadas"
                        for i_offset in range(1, 5):
                            if idx - i_offset >= 0:
                                fila_sup = df_pistas.iloc[idx - i_offset]
                                for col_idx, val in enumerate(fila_sup):
                                    val_str = str(val).strip().upper()
                                    if 'FINCA' in val_str:
                                        finca = str(fila_sup.iloc[col_idx+1]).strip() if col_idx+1 < len(fila_sup) else "N/A"
                                    if 'HECT' in val_str or 'HAS' in val_str:
                                        hectareas = str(fila_sup.iloc[col_idx+1]).strip() if col_idx+1 < len(fila_sup) else "N/A"

                        # Bajar por la lista de productos aplicados (Identificando columnas)
                        fila_encabezado = df_pistas.iloc[idx]
                        col_prod_idx, col_cant_idx, col_lote_idx = 1, 3, 4 # Por defecto
                        
                        for c_i, c_v in enumerate(fila_encabezado):
                            c_str = str(c_v).strip().upper()
                            if 'PRODUCTO' in c_str: col_prod_idx = c_i
                            elif 'CANTIDAD' in c_str or 'DOSIS' in c_str or 'TOTAL' in c_str: col_cant_idx = c_i
                            elif 'LOTE' in c_str: col_lote_idx = c_i

                        fila_actual = idx + 1
                        while fila_actual < len(df_pistas):
                            producto = str(df_pistas.iloc[fila_actual, col_prod_idx]).strip()
                            if producto.lower() == 'nan' or producto == '' or 'MEZCLA' in producto.upper() or 'TOTAL' in producto.upper():
                                break # Fin del bloque de productos
                                
                            cantidad = df_pistas.iloc[fila_actual, col_cant_idx]
                            lote = str(df_pistas.iloc[fila_actual, col_lote_idx]).strip()
                            
                            # 3. Validación con Sábana SAP (El Semáforo)
                            estado_lote = "⚠️ Validando..."
                            if col_prod_sab and col_lote_sab:
                                match_prod = df_sabana[df_sabana[col_prod_sab].astype(str).str.contains(producto, case=False, na=False, regex=False)]
                                if match_prod.empty:
                                    estado_lote = "🚨 NO EN SÁBANA"
                                else:
                                    match_lote = match_prod[match_prod[col_lote_sab].astype(str).str.contains(lote, case=False, na=False, regex=False)]
                                    if match_lote.empty:
                                        estado_lote = "❌ LOTE INVÁLIDO"
                                    else:
                                        estado_lote = "✅ LOTE OK"
                                        
                            datos_validados.append({
                                "ESTADO LOTE": estado_lote,
                                "FINCA": finca,
                                "HECTÁREAS": hectareas,
                                "PRODUCTO": producto,
                                "CANTIDAD": cantidad,
                                "LOTE PISTA": lote,
                                "ORIGEN": origen
                            })
                            fila_actual += 1
                            
                    st.session_state['df_validacion'] = pd.DataFrame(datos_validados)
                    st.success("✅ ¡Cruce Táctico Completado!")
                    
                except Exception as e:
                    st.error(f"🚨 Falla en los motores de validación: {e}")

        if 'df_validacion' in st.session_state and not st.session_state['df_validacion'].empty:
            st.markdown("### 🚦 Panel de Resultados (Pista vs Sábana)")
            def color_estado(val):
                if '✅' in str(val): return 'color: green; font-weight: bold;'
                elif '❌' in str(val) or '🚨' in str(val): return 'background-color: #ffcccc; color: red; font-weight: bold;'
                return ''
            st.dataframe(st.session_state['df_validacion'].style.map(color_estado, subset=['ESTADO LOTE']), use_container_width=True)
            # --- MOTOR DE LIQUIDACIÓN AUTOMÁTICA (BASADO EN TABLA 2) ---
            st.markdown("---")
            st.markdown("<h2 class='titulo-principal'>🧮 Motor de Liquidación de Combate</h2>", unsafe_allow_html=True)
            
            if 'df_config' in st.session_state:
                df_config = st.session_state['df_config']
                # Normalizar columnas de Config
                df_config.columns = [str(c).upper().strip() for c in df_config.columns]
                col_finca_conf = next((c for c in df_config.columns if 'FINCA' in c), None)
                col_tope_conf = next((c for c in df_config.columns if 'TOPE' in c), None)
                col_prod_conf = next((c for c in df_config.columns if 'PRODUCTOR' in c or 'TIPO' in c), None)
                
                # Buscar datos de la Finca Actual (Asumiendo que 'finca_pista' se detectó arriba)
                # NOTA: Para este ejemplo, usamos la última finca procesada o se puede hacer una selección
                
                with st.form("form_liquidacion_auto"):
                    st.info("El sistema cruzará la finca detectada con TABLA 2 para aplicar topes automáticamente.")
                    
                    c_auto1, c_auto2 = st.columns(2)
                    with c_auto1:
                        finca_seleccionada = st.text_input("Finca Detectada", value="Ej: SACRAMENTO") # Aquí inyectaremos 'finca_pista' real
                        hectareas_liq = st.number_input("Hectáreas Voladas", min_value=0.1, value=120.0, step=1.0)
                        
                        col_hor1, col_hor2 = st.columns(2)
                        hor_ini = col_hor1.number_input("Horómetro Inicial", value=0.0)
                        hor_fin = col_hor2.number_input("Horómetro Final", value=1.5)
                        valor_hora = st.number_input("Valor Hora Avión ($)", value=2500000)
                        
                    with c_auto2:
                        dias_ciclo = st.number_input("Días Ciclo / Intervalo", value=10)
                        tarifa_base_tecnico = st.number_input("Tarifa Base Serv. Técnico ($)", value=8000)
                        st.markdown("---")
                        recargo_activo = st.checkbox("⚠️ Aplica Recargo (Dom/Festivo)")
                        valor_recargo = st.number_input("Valor Recargo ($/ha)", value=5000) if recargo_activo else 0
                        
                    btn_liq_auto = st.form_submit_button("🚀 EJECUTAR CÁLCULO INTELIGENTE", type="primary", use_container_width=True)

                if btn_liq_auto:
                    # 1. BÚSQUEDA AUTOMÁTICA EN TABLA 2 (BUSCARV en Python)
                    tope_asignado = 45000 # Valor por defecto
                    tipo_productor = "ESTANDAR"
                    es_pdiv = False
                    
                    if col_finca_conf:
                        match_finca = df_config[df_config[col_finca_conf].astype(str).str.contains(finca_seleccionada.split()[0], case=False, na=False)]
                        if not match_finca.empty:
                            if col_tope_conf: tope_asignado = float(match_finca.iloc[0][col_tope_conf])
                            if col_prod_conf: tipo_productor = str(match_finca.iloc[0][col_prod_conf]).upper()
                            # Detectar si es Pista Divas por el nombre o la configuración
                            if 'PDIV' in str(match_finca.values).upper(): es_pdiv = True

                    # 2. CÁLCULOS MATEMÁTICOS DE SUS FÓRMULAS
                    tiempo_vuelo = max(0.0, hor_fin - hor_ini)
                    costo_real = (tiempo_vuelo * valor_hora) / hectareas_liq if hectareas_liq > 0 else 0
                    
                    # La Condición y la Excepción PDIV
                    if es_pdiv:
                        precio_aplicacion = costo_real + valor_recargo
                        alerta = "🟢 PDIV - Tope Ignorado"
                    else:
                        precio_aplicacion = min(costo_real, tope_asignado) + valor_recargo
                        alerta = "🔴 Tope Aplicado" if costo_real > tope_asignado else "✅ Precio Real"

                    # 3. Asistencia Técnica (Fórmula 5)
                    # Aquí puede agregar el multiplicador si un Productor A paga distinto a un Productor B
                    factor_productor = 1.0 # Si el tipo es A, 1.0, si es B, 1.2, etc. (ajustable)
                    costo_tecnico = dias_ciclo * tarifa_base_tecnico * factor_productor

                    # RESULTADOS
                    st.success(f"**Inteligencia Aplicada:** Productor: `{tipo_productor}` | Tope Detectado: `${tope_asignado:,.0f}` | Estado: `{alerta}`")
                    
                    c_res1, c_res2, c_res3 = st.columns(3)
                    c_res1.metric("⏱️ Tiempo Vuelo", f"{tiempo_vuelo:.2f} hrs")
                    c_res2.metric("🚁 TOTAL APLICACIÓN /ha", f"${precio_aplicacion:,.0f}")
                    c_res3.metric("👨‍🔧 TOTAL SERV. TÉCNICO", f"${costo_tecnico:,.0f}")
            else:
                st.info("💡 Suba la 'TABLA 2 / Configuración' en el Buzón de Carga para activar el Motor Automático.")
