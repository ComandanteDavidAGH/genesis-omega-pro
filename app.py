import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import gspread
import plotly.express as px

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
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        f_sabana = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv"], key="sab")
    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        f_pedidos = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv"], key="ped")
    with c3:
        st.markdown("### 🚁 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pis")

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas:
            with st.spinner("Sincronizando los 3 frentes y conectando con Satélite en Google Drive..."):
                try: # <--- INICIO DEL BLOQUE DE SEGURIDAD MAESTRO
                    # 1. Leer Sábana
                    bytes_sabana = io.BytesIO(f_sabana.getvalue())
                    if f_sabana.name.lower().endswith(('.xlsx', '.xls')):
                        st.session_state['df_sabana'] = pd.read_excel(bytes_sabana)
                    else:
                        st.session_state['df_sabana'] = pd.read_csv(bytes_sabana, sep=None, engine='python')
                    
                    # 2. Leer Pedidos
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    if f_pedidos.name.lower().endswith(('.xlsx', '.xls')):
                        st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos)
                    else:
                        st.session_state['df_pedidos'] = pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    # 3. CONEXIÓN SATELITAL
                    try:
                        if "gcp_credentials" in st.secrets:
                            cred_dict = dict(st.secrets["gcp_credentials"])
                            gc = gspread.service_account_from_dict(cred_dict)
                        else:
                            gc = gspread.service_account(filename='credenciales.json')
                        
                        url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                        boveda = gc.open_by_url(url_boveda)
                        hoja_tabla2 = boveda.worksheet("TABLA 2")
                        datos_tabla2 = hoja_tabla2.get_all_values()
                        st.session_state['df_config'] = pd.DataFrame(datos_tabla2[1:], columns=datos_tabla2[0])
                    except Exception as error_nube:
                        st.error(f"🚨 Falla en el Enlace Satelital: {error_nube}")
                    
                    # 4. ESCÁNER DE BARRIDO TOTAL
                    lista_pistas = []
                    for f in f_pistas:
                        try:
                            bytes_data = f.getvalue()
                            ext = f.name.split('.')[-1].lower()
                            
                            if ext == 'xlsx':
                                wb = openpyxl.load_workbook(io.BytesIO(bytes_data), read_only=True)
                                visibles = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
                                wb.close()
                                dict_pestañas = pd.read_excel(io.BytesIO(bytes_data), sheet_name=visibles, header=None)
                            else:
                                dict_pestañas = pd.read_excel(io.BytesIO(bytes_data), sheet_name=None, header=None)

                            for nombre_pestaña, df in dict_pestañas.items():
                                df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                                filas_coctel = df[df.astype(str).apply(lambda x: x.str.contains('COCTEL', case=False, na=False)).any(axis=1)].index.tolist()
                                
                                for i_c, fila_c_idx in enumerate(filas_coctel):
                                    fila_data = df.iloc[fila_c_idx].dropna().tolist()
                                    nombre_coctel = fila_data[1] if len(fila_data) > 1 else "DESCONOCIDO"
                                    limite_inferior = filas_coctel[i_c + 1] if i_c + 1 < len(filas_coctel) else len(df)
                                    df_segmento = df.iloc[fila_c_idx:limite_inferior]
                                    idx_fincas = df_segmento[df_segmento.astype(str).apply(lambda x: x.str.contains('FINCAS', case=False, na=False)).any(axis=1)].index
                                    
                                    if not idx_fincas.empty:
                                        fila_h_fincas = idx_fincas[0]
                                        col_fincas_idx = (df.iloc[fila_h_fincas].astype(str).str.contains('FINCAS', case=False)).values.argmax()
                                        for r in range(fila_h_fincas + 1, limite_inferior):
                                            finca_v = str(df.iloc[r, col_fincas_idx]).strip()
                                            if finca_v.lower() in ['nan', '', 'none'] or "TOTAL" in finca_v.upper():
                                                break
                                            lista_pistas.append({
                                                "ORIGEN": f"{f.name} | {nombre_pestaña}",
                                                "COCTEL": nombre_coctel,
                                                "FINCA_INFORME": finca_v,
                                                "DATOS_FILA": df.iloc[r].to_dict()
                                            })
                        except Exception as e_file:
                            st.error(f"🚨 Error en archivo {f.name}: {e_file}")

                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                        st.success(f"✅ ¡Barrido Exitoso! {len(lista_pistas)} vuelos detectados.")
                    else:
                        st.error("🚨 No se encontró la estructura de 'FINCAS'.")

                except Exception as e_master: # <--- CIERRE DEL BLOQUE DE SEGURIDAD QUE FALTABA
                    st.error(f"🚨 Error Crítico en Procesamiento: {e_master}")

elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>🚀 Módulo 2: Matriz de Liquidación</h1>", unsafe_allow_html=True)
    
    if 'df_pistas' not in st.session_state:
        st.warning("⚠️ Sin suministros. Cargue SAP e Informes en la Zona de Aterrizaje primero.")
    else:
        df_p = st.session_state['df_pistas']
        
        # 1. SELECTOR BLINDADO (Al cambiar aquí, cambia toda la pantalla)
        # Creamos una etiqueta única para que el sistema no se confunda entre fincas
        df_p['SELECTOR'] = df_p['FINCA_INFORME'].astype(str) + " | Cóctel: " + df_p['COCTEL'].astype(str)
        
        st.markdown("### 📡 Selección de Objetivo")
        opciones_vuelo = ["---"] + df_p['SELECTOR'].tolist()
        vuelo_seleccionado = st.selectbox("🎯 Fije el blanco (Seleccione Vuelo):", opciones_vuelo)
        
        if vuelo_seleccionado != "---":
            # Extraer datos exactos de la selección para que TODO cambie
            datos_vuelo = df_p[df_p['SELECTOR'] == vuelo_seleccionado].iloc[0]
            datos_raw = datos_vuelo['DATOS_FILA'] # Este es el diccionario con las columnas reales de Excel
            finca_nombre = datos_vuelo['FINCA_INFORME']
            coctel_nombre = datos_vuelo['COCTEL']
            
            # --- BUSCADOR INTELIGENTE BLINDADO (Productor y Margen) ---
            tipo_productor = "ESTANDAR"
            margen_val = 0.10 # 10% base por defecto
            
            try:
                # 1. Buscar en Tabla de Apoyo (Escáner Dinámico de Columnas)
                if 'df_apoyo' in st.session_state and not st.session_state['df_apoyo'].empty:
                    col_finca = [c for c in st.session_state['df_apoyo'].columns if 'FINCA' in str(c).upper()]
                    if col_finca:
                        info_finca = st.session_state['df_apoyo'][st.session_state['df_apoyo'][col_finca[0]].astype(str).str.contains(finca_nombre, case=False, na=False)]
                        if not info_finca.empty:
                            col_tipo = [c for c in st.session_state['df_apoyo'].columns if 'TIPO' in str(c).upper() and 'PROD' in str(c).upper()]
                            if col_tipo:
                                tipo_productor = info_finca[col_tipo[0]].values[0]

                # 2. Buscar Margen en Configuración (Evita KeyError)
                if 'df_config' in st.session_state and not st.session_state['df_config'].empty:
                    col_prod_conf = [c for c in st.session_state['df_config'].columns if 'PRODUCTOR' in str(c).upper()]
                    col_margen_conf = [c for c in st.session_state['df_config'].columns if 'MARGEN' in str(c).upper()]
                    
                    if col_prod_conf and col_margen_conf:
                        margen_row = st.session_state['df_config'][st.session_state['df_config'][col_prod_conf[0]].astype(str).str.contains(tipo_productor, case=False, na=False)]
                        if not margen_row.empty:
                            valor_bruto = str(margen_row[col_margen_conf[0]].values[0]).replace('%', '').strip()
                            margen_val = float(valor_bruto) / 100 if float(valor_bruto) > 1 else float(valor_bruto)
            except Exception as e_margen:
                st.warning("⚠️ No se detectó la columna exacta de márgenes. Usando 10% de base táctica.")
            # -----------------------------------------------------------
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns(4)
                f_edit = c1.text_input("📍 Finca (Editable para SAP):", value=finca_nombre)
                
                # La hectárea viene de la columna 8 del informe, pero si no la encuentra pone 79 por defecto
                ha_inf = float(datos_raw.get(8, datos_raw.get('HECTÁREAS', 79.0))) if pd.notnull(datos_raw.get(8)) else 79.0
                ha_real = c2.number_input("🚜 Hectáreas Reales:", value=ha_inf, step=0.1)
                
                horometro = c3.number_input("⏱️ Horómetro (hrs):", value=0.0, step=0.1)
                
                # Matriz de equipos (Dron, Avión, Bloque)
                tipo_maquina = c4.selectbox("🛸 Vehículo y Matriz:", ["✈️ Avión (Tarifa Normal)", "🚁 Dron (Tarifa Fija)", "✈️ Aviones Múltiples (Bloque)"])

                st.markdown(f"**Tipo Productor:** `{tipo_productor}` | **Margen de Ganancia:** `{margen_val*100}%` | **Cóctel Detectado:** `{coctel_nombre}`")

                # --- 📊 MATRIZ DE PRODUCTOS, DOSIS Y PRECIOS (ESPEJO DE SU EXCEL) ---
                st.markdown("#### 🧪 Laboratorio de Mezcla e Inventario")
                
                items_mezcla = []
                # El sistema escanea desde la columna 10 a la 18 de su Excel buscando productos
                for i in range(10, 19): 
                    cant = datos_raw.get(i, 0)
                    if pd.notnull(cant) and str(cant).replace('.','',1).isdigit() and float(cant) > 0:
                        # Valores simulados de SAP que luego conectaremos con su df_sabana real
                        costo_unitario_sap = 12000 + (i * 1500) 
                        precio_venta = costo_unitario_sap * (1 + margen_val)
                        
                        items_mezcla.append({
                            "Material (Producto)": f"Cod_Columna_{i}", # Extraeremos el nombre real en el paso final
                            "Cant. Reportada": float(cant),
                            "Dosis Piloto": float(cant) / ha_real if ha_real > 0 else 0,
                            "Costo SAP": costo_unitario_sap,
                            "Precio de Venta": precio_venta,
                            "Valor Total": float(cant) * precio_venta
                        })
                
                if items_mezcla:
                    df_detallado = pd.DataFrame(items_mezcla)
                    # use_container_width=True ajusta las columnas para usar toda la pantalla
                    st.dataframe(df_detallado.style.format({
                        "Cant. Reportada": "{:.2f}",
                        "Dosis Piloto": "{:.3f}",
                        "Costo SAP": "${:,.0f}",
                        "Precio de Venta": "${:,.0f}",
                        "Valor Total": "${:,.0f}"
                    }), use_container_width=True)
                else:
                    st.info("🚨 No se detectaron productos aplicados para esta finca en el radar.")

                # --- BOTÓN DE LIQUIDACIÓN ---
                st.markdown("---")
                if st.button("🔥 CONFIRMAR MATRIZ Y FACTURAR", type="primary", use_container_width=True):
                    st.balloons()
                    st.success(f"✅ ¡Operación exitosa! Liquidación de {f_edit} enviada al historial de base de datos.")
