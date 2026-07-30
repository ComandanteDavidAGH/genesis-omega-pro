import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# --- 🔌 CONEXIÓN ---
@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception as e:
        return None

def procesar_fecha_estricta(val):
    if pd.isna(val) or str(val).strip() == "": return pd.NaT
    s = str(val).strip()
    if s.replace('.', '', 1).isdigit(): 
        return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y'):
        try: return pd.to_datetime(s, format=fmt)
        except: pass
    return pd.NaT

# --- 🧠 MOTOR DE EXTRACCIÓN DE PRECIOS DEL ÚLTIMO AÑO ---
@st.cache_data(show_spinner=False, ttl=3600)
def extraer_catalogo_precios_reciente():
    gc = inicializar_cliente_gspread()
    if not gc: return []
    try:
        sh_precios = gc.open_by_url("https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit")
        productos_recientes = set()
        anio_actual = datetime.now().year
        
        for ws in sh_precios.worksheets():
            datos = ws.get_all_values()
            if not datos: continue
            
            idx_anio, idx_prod = -1, -1
            for i in range(min(10, len(datos))):
                fila_up = [str(x).upper().strip() for x in datos[i]]
                if 'AÑO' in fila_up and 'PRODUCTO' in fila_up:
                    idx_anio = fila_up.index('AÑO')
                    idx_prod = fila_up.index('PRODUCTO')
                    break
            
            if idx_anio != -1 and idx_prod != -1:
                for row in datos[idx_anio+1:]:
                    if len(row) > max(idx_anio, idx_prod):
                        anio_str = str(row[idx_anio]).strip()
                        if anio_str in [str(anio_actual), str(anio_actual - 1)]:
                            p_nombre = str(row[idx_prod]).strip().upper()
                            if p_nombre and "DOSIS" not in p_nombre and "SIGLAS" not in p_nombre and len(p_nombre) > 3:
                                productos_recientes.add(p_nombre)
        return list(productos_recientes)
    except Exception:
        return []

# --- DICCIONARIO BASE ---
DICT_BASE_PRODUCTOS = {
    "ACEITE DICAM": "ROYAL BIOCHEM",
    "ACONDICIONADOR SV": "SYS TECNOLOGIES",
    "ADHERENTE SV": "SYS TECNOLOGIES",
    "BANADAK": "PLANDAK",
    "BANANO Y PLATANO * LT": "INVESA S.A.S.",
    "BANATREL SC": "YARA S.A.S.",
    "BOSCALID 50 WG": "DVA COLOMBIA",
    "CERAQUINT SP": "CERADIS COLOMBIA",
    "CEROSTRESS SV * LT": "MICROFERTIZA",
    "COMPER SV": "ADAMA",
    "EPOXICONAZOLE DEL MONTE": "DEL MONTE SAS",
    "FENTRIUPH AGRO 88 OL": "DEL MONTE SAS",
    "FOSFOSTRESS SV": "MICROFERTIZA",
    "GLOBAFOL nf": "SYNGENTA",
    "IMBIOSIL O": "INBIOMA",
    "KURDO 250 EC": "INVESA S.A.S.",
    "KYVENTIQ": "CORTEVA",
    "LONSELOR 30 SC": "BASF QUÍMICA",
    "MANCOL 430 SC": "CASAGRO",
    "NATURAMIN WSP": "AGRIANDES DAINSA",
    "OPORTO": "ADAMA",
    "OPUS 12 EC": "BASF QUÍMICA",
    "POLYTHION SC": "UPL",
    "POWMYL SV": "SUMITOMO",
    "QUELAMIX": "INGEPLANT",
    "REFLECT": "SYNGENTA",
    "ROUTINE SC": "BAYER",
    "SEEKER": "SYNGENTA",
    "SICO": "SYNGENTA",
    "SIGANEX 60 SC": "BAYER",
    "SPRAYFIX": "AGRIANDES DAINSA",
    "THIOPRON 825 SC": "UPL",
    "TIMOREX PRO": "ADAMA",
    "XILOTROM": "AGRIFOL",
    "ZINTRAC x LITRO SV": "YARA S.A.S."
}

# --- 🚀 EJECUCIÓN DEL MÓDULO ---
def ejecutar():
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    
    /* 💥 KPIs PERSONALIZADOS Y BLINDADOS */
    .kpi-card { background-color: #0d1b2a; color: white; padding: 20px; border-radius: 10px; border-left: 6px solid #d4af37; box-shadow: 0 4px 6px rgba(0,0,0,0.2); margin-bottom: 15px; }
    .kpi-rojo { border-left-color: #dc3545; }
    .kpi-amarillo { border-left-color: #ffc107; }
    .kpi-verde { border-left-color: #28a745; }
    .kpi-titulo { font-weight: bold; font-size: 14px; margin-bottom: 5px; text-transform: uppercase; color: #a0aec0; }
    .kpi-valor { font-size: 28px; font-weight: 900; margin: 0; color: white; }

    /* 💥 FORZAR ETIQUETAS E INPUTS */
    div[data-testid="stMainBlockContainer"] label p { 
        color: #0d1b2a !important; 
        font-weight: 900 !important; 
        text-transform: uppercase !important; 
        font-size: 13px !important; 
    }
    
    div[data-testid="stTextInput"] input, 
    div[data-testid="stNumberInput"] input, 
    div[data-testid="stDateInput"] input {
        border: 2px solid #0d1b2a !important; 
        border-radius: 6px !important; 
        color: #000000 !important; 
        font-weight: 900 !important; 
        background-color: #ffffff !important;
    }
    
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        border: 2px solid #0d1b2a !important; 
        border-radius: 6px !important; 
        background-color: #ffffff !important;
    }
    
    div[data-testid="stSelectbox"] div[data-baseweb="select"] * { 
        color: #000000 !important; 
        font-weight: 900 !important; 
    }
    
    .st-expander { border: 2px solid #0d1b2a !important; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

    URL_SHEET_LOCAL = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"

    c_tit, c_btn1, c_btn2 = st.columns([2, 1, 1])
    c_tit.markdown("<h1 class='titulo-mod'>📦 19. Control y Auditoría</h1>", unsafe_allow_html=True)
    
    if c_btn1.button("🔄 REFRESCAR RADARES", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    if c_btn2.button("🧹 PURGAR DICCIONARIO", type="secondary", use_container_width=True):
        with st.spinner("Aniquilando diccionario oculto..."):
            gc = inicializar_cliente_gspread()
            if gc:
                try:
                    sh_local = gc.open_by_url(URL_SHEET_LOCAL)
                    ws_dicc = sh_local.worksheet("DICCIONARIO")
                    sh_local.del_worksheet(ws_dicc)
                    st.cache_data.clear()
                    st.success("✅ ¡Fantasma destruido! La lista ha sido purgada.")
                    st.rerun()
                except Exception:
                    st.info("No había basura. La lista ya estaba limpia.")

    st.write("Panel táctico de auditoría. Ingresa lotes cruzando información oficial con la Base de Precios SAP.")

    gc = inicializar_cliente_gspread()
    if not gc:
        st.error("🚨 Servidor desconectado. Revisa tus credenciales de Google Cloud.")
        return

    with st.spinner("📡 Sincronizando Bóveda de Ingresos y Catálogo de Precios Históricos..."):
        try:
            sh_local = gc.open_by_url(URL_SHEET_LOCAL)
            ws_ingresos = sh_local.get_worksheet(0) 
            datos_crudos = ws_ingresos.get_all_values()
            
            # --- FUSIÓN TÁCTICA ---
            dict_operativo = {k.upper(): v.upper() for k, v in DICT_BASE_PRODUCTOS.items()}
            
            try:
                ws_dicc = sh_local.worksheet("DICCIONARIO")
                datos_dicc = ws_dicc.get_all_values()
                for row in datos_dicc[1:]: 
                    if len(row) >= 2 and str(row[0]).strip():
                        producto_nube = str(row[0]).strip().upper()
                        proveedor_nube = str(row[1]).strip().upper()
                        if len(producto_nube) > 3: 
                            dict_operativo[producto_nube] = proveedor_nube
            except Exception:
                pass 
            
            productos_precio_sap = extraer_catalogo_precios_reciente()
            for prod_sap in productos_precio_sap:
                if prod_sap not in dict_operativo:
                    dict_operativo[prod_sap] = "" 

        except Exception as e:
            st.error(f"🚨 Error de acceso a las Bóvedas. Detalle: {e}")
            return

    if not datos_crudos or len(datos_crudos) < 2:
        st.warning("La base de datos parece estar vacía.")
        return

    idx_header = 0
    for i, row in enumerate(datos_crudos[:5]):
        fila_upper = [str(x).upper().strip() for x in row]
        if "ESTADO / OBSERVACIÓN" in fila_upper or "PRODUCTO" in fila_upper:
            idx_header = i
            break

    encabezados = [str(x).strip().upper() for x in datos_crudos[idx_header]]
    df = pd.DataFrame(datos_crudos[idx_header+1:], columns=encabezados)
    
    col_producto = next((c for c in df.columns if "PRODUCTO" in c), None)
    if col_producto:
        df = df[df[col_producto].str.strip() != ""]

    df['FILA_EXCEL'] = range(idx_header + 2, len(df) + idx_header + 2)

    COL_ESTADO = "ESTADO / OBSERVACIÓN"
    if COL_ESTADO not in df.columns:
        st.error(f"🚨 FALTA COLUMNA TÁCTICA: No se encontró la columna **{COL_ESTADO}**.")
        return

    idx_col_estado = encabezados.index(COL_ESTADO) + 1 

    # --- 📊 1. PANEL DE RADARES (KPIs) ---
    st.markdown("### 📡 Radares de Vencimiento")
    
    hoy = datetime.now()
    limite_90_dias = hoy + timedelta(days=90)
    
    col_fv = next((c for c in df.columns if c in ["F/V", "FECHA VENCIMIENTO", "VENCIMIENTO"]), None)
    
    lotes_vencidos = 0
    lotes_riesgo = 0
    
    if col_fv:
        df['FECHA_VENC_DT'] = df[col_fv].apply(procesar_fecha_estricta)
        df_activos = df[~df[COL_ESTADO].str.contains("ANULADO", na=False, case=False)]
        lotes_vencidos = df_activos[df_activos['FECHA_VENC_DT'] < hoy].shape[0]
        lotes_riesgo = df_activos[(df_activos['FECHA_VENC_DT'] >= hoy) & (df_activos['FECHA_VENC_DT'] <= limite_90_dias)].shape[0]

    k1, k2, k3 = st.columns(3)
    k1.markdown(f"""
    <div class='kpi-card kpi-verde'>
        <div class='kpi-titulo'>📦 Ingresos Registrados</div>
        <p class='kpi-valor'>{len(df)}</p>
    </div>
    """, unsafe_allow_html=True)

    k2.markdown(f"""
    <div class='kpi-card kpi-rojo'>
        <div class='kpi-titulo'>🚨 Lotes Vencidos</div>
        <p class='kpi-valor'>{lotes_vencidos}</p>
    </div>
    """, unsafe_allow_html=True)

    k3.markdown(f"""
    <div class='kpi-card kpi-amarillo'>
        <div class='kpi-titulo'>⚠️ Por Vencer (90 Días)</div>
        <p class='kpi-valor'>{lotes_riesgo}</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # --- ➕ FORMULARIO DINÁMICO DE INYECCIÓN ---
    st.markdown("### ➕ Inyector de Nuevos Ingresos")
    
    with st.container(border=True):
        st.markdown("<p style='color: #0d1b2a; font-size: 16px; font-weight: 900;'>1. IDENTIFICACIÓN DEL QUÍMICO OFICIAL</p>", unsafe_allow_html=True)
        
        c_tog1, c_tog2 = st.columns(2)
        es_nuevo_producto = c_tog1.toggle("✨ Ingresar un Producto Totalmente NUEVO")
        modificar_prov = False
        
        c_prod, c_prov = st.columns(2)
        
        if es_nuevo_producto:
            n_prod = c_prod.text_input("Nombre del Nuevo Producto")
            n_prov = c_prov.text_input("Nombre del Proveedor")
        else:
            modificar_prov = c_tog2.toggle("✏️ Corregir / Modificar Proveedor")
            lista_prods_ordenada = sorted([p for p in dict_operativo.keys() if len(p) > 3])
            
            n_prod = c_prod.selectbox("Seleccione el Producto (Integrado con Precios SAP)", lista_prods_ordenada)
            proveedor_asignado = dict_operativo.get(n_prod, "")
            
            es_vacio = not bool(proveedor_asignado.strip())
            debe_desbloquear = modificar_prov or es_vacio
            
            n_prov = c_prov.text_input(
                "Proveedor", 
                value=proveedor_asignado, 
                disabled=not debe_desbloquear, 
                placeholder="Digite el proveedor para guardarlo en el Diccionario"
            )
            
        st.markdown("<p style='color: #0d1b2a; font-size: 16px; font-weight: 900; margin-top: 15px;'>2. DATOS OPERATIVOS</p>", unsafe_allow_html=True)
        f1, f2, f3 = st.columns(3)
        
        n_fecha_ing = f2.date_input("Fecha de Ingreso a SAP")
        semana_calculada = n_fecha_ing.isocalendar()[1]
        n_semana = f1.text_input("Semana del Año (Automática)", value=str(semana_calculada), disabled=True)
        
        n_pista = f3.selectbox("Almacén SAP (Pista)", ["LUCI", "PLUC", "PDIV", "PORI", "TEHO"])
        
        f4, f5, f6 = st.columns(3)
        n_cant = f4.number_input("Cantidad", min_value=0.0, step=1.0)
        n_lote = f5.text_input("Lote")
        n_ff = f6.date_input("Fecha de Fabricación (F/F)")
        
        f7, f8, f9, f10 = st.columns(4)
        n_fv = f7.date_input("Fecha de Vencimiento (F/V)")
        n_factura = f8.text_input("Factura")
        n_pedido = f9.text_input("Pedido")
        n_consecutivo = f10.text_input("Consecutivo (SAP)")
        
        btn_guardar_nuevo = st.button("🚀 INYECTAR NUEVO LOTE A LA BÓVEDA", type="primary", use_container_width=True)
        
        if btn_guardar_nuevo:
            if not n_prod or str(n_prod).strip() == "":
                st.error("🚨 El nombre del producto no puede estar vacío.")
            else:
                prod_limpio = str(n_prod).strip().upper()
                prov_limpio = str(n_prov).strip().upper()

                actualizar_dicc = False
                if es_nuevo_producto:
                    actualizar_dicc = True
                elif (modificar_prov or es_vacio) and prov_limpio and prov_limpio != proveedor_asignado.upper():
                    actualizar_dicc = True
                
                if actualizar_dicc:
                    try:
                        ws_dicc = sh_local.worksheet("DICCIONARIO")
                    except Exception:
                        ws_dicc = sh_local.add_worksheet(title="DICCIONARIO", rows="100", cols="2")
                        ws_dicc.append_row(["PRODUCTO", "PROVEEDOR"])
                    
                    try:
                        datos_d = ws_dicc.get_all_values()
                        fila_a_actualizar = -1
                        for idx_d, row_d in enumerate(datos_d):
                            if len(row_d) > 0 and str(row_d[0]).strip().upper() == prod_limpio:
                                fila_a_actualizar = idx_d + 1
                                break
                        
                        if fila_a_actualizar != -1:
                            ws_dicc.update_cell(fila_a_actualizar, 2, prov_limpio)
                        else:
                            ws_dicc.append_row([prod_limpio, prov_limpio])
                    except Exception as e:
                        st.warning(f"Se guardó el ingreso, pero falló la escritura en el Diccionario: {e}")

                nueva_fila_drive = []
                for header in encabezados:
                    h = header.upper()
                    if "SEMANA" in h: nueva_fila_drive.append(str(semana_calculada))
                    elif "PROVEEDOR" in h: nueva_fila_drive.append(prov_limpio)
                    elif "FECHA DE INGRESO" in h: nueva_fila_drive.append(n_fecha_ing.strftime("%d/%m/%Y"))
                    elif "PRODUCTO" in h: nueva_fila_drive.append(prod_limpio)
                    elif "PISTA" in h: nueva_fila_drive.append(str(n_pista))
                    elif "CANTIDAD" in h: nueva_fila_drive.append(str(n_cant))
                    elif "LOTE" in h: nueva_fila_drive.append(str(n_lote))
                    elif "F/F" in h: nueva_fila_drive.append(n_ff.strftime("%d/%m/%Y"))
                    elif "F/V" in h: nueva_fila_drive.append(n_fv.strftime("%d/%m/%Y"))
                    elif "FACTURA" in h: nueva_fila_drive.append(str(n_factura))
                    elif "PEDIDO" in h: nueva_fila_drive.append(str(n_pedido))
                    elif "CONSECUTIVO" in h: nueva_fila_drive.append(str(n_consecutivo))
                    elif "ESTADO" in h: nueva_fila_drive.append("✅ VIGENTE")
                    else: nueva_fila_drive.append("") 
                
                try:
                    with st.spinner("Enviando datos al satélite..."):
                        ws_ingresos.append_row(nueva_fila_drive)
                    st.success(f"✅ ¡Lote de {prod_limpio} inyectado exitosamente en SAP-Nube!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al inyectar datos: {e}")

    # --- 📧 GENERADOR DE REPORTE PARA CORREO (CONSTRUCTOR HTML MANUAL EXTREMO) ---
    st.markdown("---")
    st.markdown("### 📧 Reporte Rápido para Correo (Copy & Paste)")
    st.info("💡 Selecciona la fecha de los ingresos. Sombrea la tabla resultante, cópiala y pégala directo en tu correo. El formato se conservará impecable.")
    
    col_fecha_rep, col_vacia = st.columns([1, 3])
    fecha_reporte = col_fecha_rep.date_input("Fecha a reportar:", value=datetime.now())
    fecha_reporte_str = fecha_reporte.strftime("%d/%m/%Y")
    
    col_fecha_ingreso = next((c for c in df.columns if "FECHA DE INGRESO" in c), None)
    
    if col_fecha_ingreso:
        df_correo = df[df[col_fecha_ingreso] == fecha_reporte_str].copy()
        
        if not df_correo.empty:
            cols_deseadas = [c for c in df.columns if c in ["SEMANA", "PROVEEDOR", "FECHA DE INGRESO", "PRODUCTO", "PISTA", "CANTIDAD", "LOTE", "F/F", "F/V", "FACTURA", "PEDIDO", "CONSECUTIVO"]]
            df_correo = df_correo[cols_deseadas]
            
            # 💥 CIRUGÍA MAYOR: CONSTRUCCIÓN MANUAL DE CADA ETIQUETA HTML PARA BYPASS A STREAMLIT
            html_manual = """
            <table style='border-collapse: collapse; width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 13px; border: 2px solid #0d1b2a; margin-top: 10px; background-color: #ffffff;'>
                <thead>
                    <tr>
            """
            # Inyectar Encabezados
            for col_name in df_correo.columns:
                html_manual += f"<th style='background-color: #0d1b2a; color: #d4af37; padding: 12px 10px; border: 2px solid #0d1b2a; text-align: center; font-weight: 900; text-transform: uppercase;'>{col_name}</th>"
            
            html_manual += """
                    </tr>
                </thead>
                <tbody>
            """
            # Inyectar Celdas de Datos
            for _, row in df_correo.iterrows():
                html_manual += "<tr>"
                for col_name in df_correo.columns:
                    val = str(row[col_name]) if pd.notna(row[col_name]) else ""
                    html_manual += f"<td style='padding: 10px; border: 2px solid #0d1b2a; text-align: center; color: #000000; font-weight: bold;'>{val}</td>"
                html_manual += "</tr>"
                
            html_manual += """
                </tbody>
            </table>
            """
            
            # Renderizar en pantalla
            st.markdown(html_manual, unsafe_allow_html=True)
            
        else:
            st.warning(f"No se encontraron ingresos registrados en la bóveda con la fecha {fecha_reporte_str}.")

    # --- 🔍 FILTROS TÁCTICOS ---
    st.markdown("---")
    st.markdown("### 🔍 Escáner de Auditoría (Filtros)")
    filtro_seleccionado = st.radio("Mostrar ingresos:", 
                                 ["🌐 Mostrar Todos", "✅ Solo Vigentes", "🚨 Solo Vencidos", "⚠️ Por Vencer (90 Días)"], 
                                 horizontal=True)
    
    df[COL_ESTADO] = df[COL_ESTADO].replace(r'^\s*$', '✅ VIGENTE', regex=True).fillna('✅ VIGENTE')

    df_filtrado = df.copy()
    if filtro_seleccionado == "✅ Solo Vigentes":
        df_filtrado = df_filtrado[df_filtrado[COL_ESTADO].str.contains("VIGENTE", case=False, na=False)]
    elif filtro_seleccionado == "🚨 Solo Vencidos" and col_fv:
        df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO", na=False)) & (df_filtrado['FECHA_VENC_DT'] < hoy)]
    elif filtro_seleccionado == "⚠️ Por Vencer (90 Días)" and col_fv:
        df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO", na=False)) & (df_filtrado['FECHA_VENC_DT'] >= hoy) & (df_filtrado['FECHA_VENC_DT'] <= limite_90_dias)]

    # --- 🛠️ TABLA DE ANULACIONES ---
    st.markdown("### 🛠️ Matriz de Anulaciones (Solo Lectura y Edición de Estado)")
    st.caption("🔒 Las cantidades y fechas están bloqueadas por seguridad. Solo puedes hacer doble clic en ESTADO/OBSERVACIÓN para anular.")
    
    cols_disabled = [col for col in df_filtrado.columns if col not in [COL_ESTADO, 'FILA_EXCEL', 'FECHA_VENC_DT']]
    
    opciones_estado = [
        "✅ VIGENTE",
        "❌ ANULADO: ERROR EN PRECIOS",
        "❌ ANULADO: ERROR DE CANTIDAD",
        "❌ ANULADO: DEVOLUCIÓN A PROVEEDOR",
        "❌ ANULADO: ERROR EN LOTE/FECHAS",
        "❌ ANULADO: OTRO MOTIVO"
    ]

    columnas_vista = [c for c in df_filtrado.columns if c not in ['FILA_EXCEL', 'FECHA_VENC_DT']]
    df_vista = df_filtrado[columnas_vista].copy()

    df_editado = st.data_editor(
        df_vista,
        column_config={
            COL_ESTADO: st.column_config.SelectboxColumn(
                "ESTADO / OBSERVACIÓN",
                help="Doble clic para anular o cambiar estado.",
                width="large",
                options=opciones_estado,
                required=True
            )
        },
        disabled=cols_disabled,
        hide_index=True,
        use_container_width=True,
        key="editor_ingresos"
    )

    # --- 💾 3. MOTOR DE SINCRONIZACIÓN ---
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 SINCRONIZAR ANULACIONES EN DRIVE", type="primary"):
        cambios_detectados = False
        
        for i in range(len(df_filtrado)):
            estado_original = str(df_filtrado.iloc[i][COL_ESTADO]).strip()
            estado_nuevo = str(df_editado.iloc[i][COL_ESTADO]).strip()
            
            if estado_original != estado_nuevo:
                fila_excel = df_filtrado.iloc[i]['FILA_EXCEL']
                with st.spinner(f"Inyectando Anulación en Fila {fila_excel}..."):
                    try:
                        ws_ingresos.update_cell(fila_excel, idx_col_estado, estado_nuevo)
                        cambios_detectados = True
                    except Exception as e:
                        st.error(f"Error al actualizar fila {fila_excel}. Detalle: {e}")
        
        if cambios_detectados:
            st.success("✅ ¡Misión Cumplida! Bóveda de Drive actualizada exitosamente con las anulaciones.")
            st.cache_data.clear() 
            st.rerun()
        else:
            st.info("No se detectaron cambios en los estados.")

    # --- 📥 DESCARGA A EXCEL ---
    st.markdown("---")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_editado.to_excel(writer, sheet_name='Auditoria_Ingresos', index=False)
        ws_excel = writer.sheets['Auditoria_Ingresos']
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        for cell in ws_excel[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for col in ws_excel.columns:
            max_length = 0
            column = col[0].column_letter 
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(cell.value)
                except: pass
            adjusted_width = (max_length + 2)
            ws_excel.column_dimensions[column].width = adjusted_width

    st.download_button(
        label="💾 DESCARGAR REPORTE DE AUDITORÍA (EXCEL)",
        data=buffer.getvalue(),
        file_name=f"Reporte_Auditoria_Ingresos_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if __name__ == "__main__":
    ejecutar()
