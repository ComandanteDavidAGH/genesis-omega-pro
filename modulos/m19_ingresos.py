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

# --- 🚀 EJECUCIÓN DEL MÓDULO ---
def ejecutar():
    st.markdown("""
    <style>
    .titulo-mod { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }
    div[data-testid="metric-container"] { background-color: #0d1b2a; border: 2px solid #d4af37; border-radius: 8px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    div[data-testid="metric-container"] label { color: #a0aec0 !important; font-weight: bold !important; font-size: 14px !important; text-transform: uppercase; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #ffffff !important; font-weight: 900 !important; font-size: 32px !important; }
    </style>
    """, unsafe_allow_html=True)

    c_tit, c_btn = st.columns([3, 1])
    c_tit.markdown("<h1 class='titulo-mod'>📦 19. Control y Auditoría de Ingresos</h1>", unsafe_allow_html=True)
    
    if c_btn.button("🔄 REFRESCAR RADARES", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.write("Panel táctico de auditoría. Identifica vencimientos y ejecuta anulaciones de ingresos directo a la base maestra.")

    gc = inicializar_cliente_gspread()
    if not gc:
        st.error("🚨 Servidor desconectado. Revisa tus credenciales de Google Cloud.")
        return

    URL_SHEET = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"

    with st.spinner("📡 Escaneando Bóveda de Ingresos en Drive..."):
        try:
            sh = gc.open_by_url(URL_SHEET)
            ws = sh.get_worksheet(0) 
            datos_crudos = ws.get_all_values()
        except Exception as e:
            st.error(f"🚨 Error de acceso. Detalle: {e}")
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
    
    # 💥 FILTRO PURIFICADOR: Destruimos filas fantasma (donde el producto está en blanco)
    col_producto = next((c for c in df.columns if "PRODUCTO" in c), None)
    if col_producto:
        df = df[df[col_producto].str.strip() != ""]

    df['FILA_EXCEL'] = range(idx_header + 2, len(df) + idx_header + 2)

    COL_ESTADO = "ESTADO / OBSERVACIÓN"
    if COL_ESTADO not in df.columns:
        st.error(f"🚨 FALTA COLUMNA TÁCTICA: No se encontró la columna **{COL_ESTADO}**.")
        return

    idx_col_estado = encabezados.index(COL_ESTADO) + 1 

    # --- 📊 1. PANEL DE RADARES (KPIs BLINDADOS) ---
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

    # --- ➕ FORMULARIO DE INYECCIÓN DE DATOS ---
    with st.expander("➕ REGISTRAR NUEVO INGRESO (ENVIAR A DRIVE)", expanded=False):
        st.info("Al guardar, este ingreso se inyectará automáticamente como una nueva fila en Google Sheets.")
        with st.form("form_nuevo_ingreso"):
            f1, f2, f3 = st.columns(3)
            n_semana = f1.text_input("Semana del Año")
            n_prov = f2.text_input("Proveedor")
            n_fecha_ing = f3.date_input("Fecha de Ingreso")
            
            f4, f5, f6 = st.columns(3)
            n_prod = f4.text_input("Producto")
            n_pista = f5.selectbox("Pista", ["LUCHA", "ORIHUCA", "PORI", "PLUC", "TEHO", "PDIV"])
            n_cant = f6.number_input("Cantidad", min_value=0.0, step=1.0)
            
            f7, f8, f9 = st.columns(3)
            n_lote = f7.text_input("Lote")
            n_ff = f8.date_input("Fecha de Fabricación (F/F)")
            n_fv = f9.date_input("Fecha de Vencimiento (F/V)")
            
            f10, f11, f12 = st.columns(3)
            n_factura = f10.text_input("Factura")
            n_pedido = f11.text_input("Pedido")
            n_consecutivo = f12.text_input("Consecutivo")
            
            btn_guardar_nuevo = st.form_submit_button("🚀 INYECTAR A LA BÓVEDA SAP", use_container_width=True)
            
            if btn_guardar_nuevo:
                if not n_prod:
                    st.error("El nombre del producto es obligatorio.")
                else:
                    nueva_fila_drive = []
                    for header in encabezados:
                        h = header.upper()
                        if "SEMANA" in h: nueva_fila_drive.append(str(n_semana))
                        elif "PROVEEDOR" in h: nueva_fila_drive.append(str(n_prov))
                        elif "FECHA DE INGRESO" in h: nueva_fila_drive.append(n_fecha_ing.strftime("%d/%m/%Y"))
                        elif "PRODUCTO" in h: nueva_fila_drive.append(str(n_prod).upper())
                        elif "PISTA" in h: nueva_fila_drive.append(str(n_pista))
                        elif "CANTIDAD" in h: nueva_fila_drive.append(str(n_cant))
                        elif "LOTE" in h: nueva_fila_drive.append(str(n_lote))
                        elif "F/F" in h: nueva_fila_drive.append(n_ff.strftime("%d/%m/%Y"))
                        elif "F/V" in h: nueva_fila_drive.append(n_fv.strftime("%d/%m/%Y"))
                        elif "FACTURA" in h: nueva_fila_drive.append(str(n_factura))
                        elif "PEDIDO" in h: nueva_fila_drive.append(str(n_pedido))
                        elif "CONSECUTIVO" in h: nueva_fila_drive.append(str(n_consecutivo))
                        elif "ESTADO" in h: nueva_fila_drive.append("✅ VIGENTE")
                        else: nueva_fila_drive.append("") # Para columnas vacías extra
                    
                    try:
                        with st.spinner("Enviando datos al satélite..."):
                            ws.append_row(nueva_fila_drive)
                        st.success("✅ Ingreso registrado con éxito.")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al guardar: {e}")

    # --- 🔍 FILTROS TÁCTICOS ---
    st.markdown("### 🔍 Filtro de Escaneo")
    filtro_seleccionado = st.radio("Mostrar ingresos:", 
                                 ["🌐 Mostrar Todos", "✅ Solo Vigentes", "🚨 Solo Vencidos", "⚠️ Por Vencer (90 Días)"], 
                                 horizontal=True)
    
    df[COL_ESTADO] = df[COL_ESTADO].replace(r'^\s*$', '✅ VIGENTE', regex=True).fillna('✅ VIGENTE')

    # Aplicar el filtro a la matriz
    df_filtrado = df.copy()
    if filtro_seleccionado == "✅ Solo Vigentes":
        df_filtrado = df_filtrado[df_filtrado[COL_ESTADO].str.contains("VIGENTE", case=False, na=False)]
    elif filtro_seleccionado == "🚨 Solo Vencidos" and col_fv:
        df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO", na=False)) & (df_filtrado['FECHA_VENC_DT'] < hoy)]
    elif filtro_seleccionado == "⚠️ Por Vencer (90 Días)" and col_fv:
        df_filtrado = df_filtrado[(~df_filtrado[COL_ESTADO].str.contains("ANULADO", na=False)) & (df_filtrado['FECHA_VENC_DT'] >= hoy) & (df_filtrado['FECHA_VENC_DT'] <= limite_90_dias)]

    # --- 🛠️ 2. TABLA DE AUDITORÍA Y ANULACIONES ---
    st.markdown("### 🛠️ Matriz de Auditoría")
    
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
                help="Doble clic para cambiar el estado operativo.",
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
    if st.button("💾 EJECUTAR Y SINCRONIZAR ANULACIONES EN DRIVE", type="primary"):
        cambios_detectados = False
        
        for i in range(len(df_filtrado)):
            estado_original = str(df_filtrado.iloc[i][COL_ESTADO]).strip()
            estado_nuevo = str(df_editado.iloc[i][COL_ESTADO]).strip()
            
            if estado_original != estado_nuevo:
                fila_excel = df_filtrado.iloc[i]['FILA_EXCEL']
                with st.spinner(f"Inyectando Anulación en Fila {fila_excel}..."):
                    try:
                        ws.update_cell(fila_excel, idx_col_estado, estado_nuevo)
                        cambios_detectados = True
                    except Exception as e:
                        st.error(f"Error al actualizar fila {fila_excel}. Detalle: {e}")
        
        if cambios_detectados:
            st.success("✅ ¡Misión Cumplida! Bóveda de Drive actualizada exitosamente.")
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
        
        # Formato básico
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
