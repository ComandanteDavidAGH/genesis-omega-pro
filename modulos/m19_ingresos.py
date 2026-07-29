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

# --- DICCIONARIO BASE TEMPORAL ---
DICT_BASE_PRODUCTOS = {}

# --- 🚀 EJECUCIÓN DEL MÓDULO ---
def ejecutar():
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    div[data-testid="metric-container"] { background-color: #0d1b2a; border: 2px solid #d4af37; border-radius: 8px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    div[data-testid="metric-container"] label { color: #a0aec0 !important; font-weight: bold !important; font-size: 14px !important; text-transform: uppercase; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #ffffff !important; font-weight: 900 !important; font-size: 32px !important; }
    .st-expander { border: 2px solid #0d1b2a !important; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

    c_tit, c_btn = st.columns([3, 1])
    c_tit.markdown("<h1 class='titulo-mod'>📦 19. Control y Auditoría de Ingresos</h1>", unsafe_allow_html=True)
    
    if c_btn.button("🔄 REFRESCAR RADARES", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.write("Panel táctico de auditoría. Ingresa lotes de manera asistida usando nomenclatura oficial SAP.")

    gc = inicializar_cliente_gspread()
    if not gc:
        st.error("🚨 Servidor desconectado. Revisa tus credenciales de Google Cloud.")
        return

    URL_SHEET_LOCAL = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"
    URL_MAESTRA_SAP = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

    with st.spinner("📡 Escaneando Bóveda Local y Base Maestra SAP..."):
        try:
            # 1. Traer Bóveda Local (Ingresos)
            sh_local = gc.open_by_url(URL_SHEET_LOCAL)
            ws_ingresos = sh_local.get_worksheet(0) 
            datos_crudos = ws_ingresos.get_all_values()
            
            # 2. Traer Nombres Oficiales de SAP
            sh_sap = gc.open_by_url(URL_MAESTRA_SAP)
            ws_cfg_sap = sh_sap.worksheet("Configuración")
            datos_cfg_sap = ws_cfg_sap.get_all_values()
            
            productos_oficiales_sap = set()
            idx_header_cfg = 0
            for i, row in enumerate(datos_cfg_sap[:5]):
                r_c = [str(x).upper().strip() for x in row]
                if 'PRODUCTO' in r_c:
                    idx_header_cfg = i
                    break
                    
            encabezados_cfg = [str(x).upper().strip() for x in datos_cfg_sap[idx_header_cfg]]
            if 'PRODUCTO' in encabezados_cfg:
                col_prod_idx = encabezados_cfg.index('PRODUCTO')
                for row in datos_cfg_sap[idx_header_cfg+1:]:
                    if len(row) > col_prod_idx:
                        p_sap = str(row[col_prod_idx]).strip().upper()
                        if p_sap and p_sap not in ["NAN", "NONE", "", "PRODUCTO"]:
                            productos_oficiales_sap.add(p_sap)
            
            # 3. Traer Diccionario de Proveedores
            dict_operativo = DICT_BASE_PRODUCTOS.copy()
            try:
                ws_dicc = sh_local.worksheet("DICCIONARIO")
                datos_dicc = ws_dicc.get_all_values()
                for row in datos_dicc[1:]: 
                    if len(row) >= 2 and str(row[0]).strip():
                        producto_nube = str(row[0]).strip().upper()
                        proveedor_nube = str(row[1]).strip().upper()
                        dict_operativo[producto_nube] = proveedor_nube
            except Exception:
                pass 
            
            # 4. Fusionar SAP + Diccionario (Aseguramos que todos los de SAP estén en la lista)
            for p in productos_oficiales_sap:
                if p not in dict_operativo:
                    dict_operativo[p] = "" # Agregarlo sin proveedor asignado

        except Exception as e:
            st.error(f"🚨 Error de acceso a las Bóvedas. Detalle: {e}")
            return

    if not datos_crudos or len(datos_crudos) < 2:
        st.warning("La base de datos parece estar vacía.")
        return

    # 💥 ESCÁNER DINÁMICO DE ENCABEZADOS
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
    k1.metric("📦 Ingresos Registrados", len(df))
    k2.metric("🚨 Lotes Vencidos", lotes_vencidos)
    k3.metric("⚠️ Por Vencer (90 Días)", lotes_riesgo)

    st.markdown("---")

    # --- ➕ FORMULARIO DINÁMICO DE INYECCIÓN ---
    st.markdown("### ➕ Inyector de Nuevos Ingresos (Base SAP)")
    
    with st.container(border=True):
        st.markdown("<p style='color: #0d1b2a; font-weight: bold;'>1. Identificación del Químico Oficial</p>", unsafe_allow_html=True)
        es_nuevo_producto = st.toggle("✨ Ingresar un Producto Totalmente NUEVO (Aún no existe en SAP)")
        
        c_prod, c_prov = st.columns(2)
        
        # --- LÓGICA DE AUTOLLENADO INTELIGENTE ---
        proveedor_asignado = ""
        if es_nuevo_producto:
            n_prod = c_prod.text_input("Nombre del Nuevo Producto")
            n_prov = c_prov.text_input("Nombre del Proveedor")
        else:
            lista_prods_ordenada = sorted(list(dict_operativo.keys()))
            n_prod = c_prod.selectbox("Seleccione el Producto (Catálogo SAP)", lista_prods_ordenada)
            proveedor_asignado = dict_operativo.get(n_prod, "")
            
            # 💥 SI EL PRODUCTO TIENE PROVEEDOR, LO BLOQUEA. SI NO TIENE, DEJA QUE EL USUARIO LO ESCRIBA.
            tiene_prov = bool(proveedor_asignado.strip())
            n_prov = c_prov.text_input("Proveedor", value=proveedor_asignado, disabled=tiene_prov, placeholder="Digite el proveedor para guardarlo en la BD")
            
        st.markdown("<p style='color: #0d1b2a; font-weight: bold; margin-top: 15px;'>2. Datos Operativos</p>", unsafe_allow_html=True)
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
        n_consecutivo = f10.text_input("Consecutivo")
        
        btn_guardar_nuevo = st.button("🚀 INYECTAR NUEVO LOTE A LA BÓVEDA", type="primary", use_container_width=True)
        
        if btn_guardar_nuevo:
            if not n_prod or str(n_prod).strip() == "":
                st.error("🚨 El nombre del producto no puede estar vacío.")
            else:
                prod_limpio = str(n_prod).strip().upper()
                prov_limpio = str(n_prov).strip().upper()

                # 1. Guardar o actualizar en DICCIONARIO si es nuevo o si se digitó el proveedor por primera vez
                if es_nuevo_producto or (not tiene_prov and prov_limpio):
                    try:
                        ws_dicc = sh_local.worksheet("DICCIONARIO")
                    except Exception:
                        ws_dicc = sh_local.add_worksheet(title="DICCIONARIO", rows="100", cols="2")
                        ws_dicc.append_row(["PRODUCTO", "PROVEEDOR"])
                    
                    try:
                        ws_dicc.append_row([prod_limpio, prov_limpio])
                    except Exception as e:
                        st.warning(f"Se guardó el ingreso, pero falló la escritura en el Diccionario: {e}")

                # 2. Inyectar datos a Bóveda
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
