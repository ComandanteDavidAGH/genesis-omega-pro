import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re

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
    .titulo-mod {{ color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; text-transform: uppercase; }}
    .kpi-card {{ background-color: #0d1b2a; color: white; padding: 20px; border-radius: 10px; border-left: 6px solid #d4af37; box-shadow: 0 4px 6px rgba(0,0,0,0.2); margin-bottom: 15px; }}
    .kpi-rojo {{ border-left-color: #dc3545; }}
    .kpi-amarillo {{ border-left-color: #ffc107; }}
    .kpi-verde {{ border-left-color: #28a745; }}
    .kpi-titulo {{ font-weight: bold; font-size: 14px; margin-bottom: 5px; text-transform: uppercase; color: #a0aec0; }}
    .kpi-valor {{ font-size: 28px; font-weight: 900; margin: 0; color: white; }}
    </style>
    """, unsafe_allow_html=True)

    c_tit, c_btn = st.columns([3, 1])
    c_tit.markdown("<h1 class='titulo-mod'>📦 19. Control y Auditoría de Ingresos</h1>", unsafe_allow_html=True)
    
    if c_btn.button("🔄 REFRESCAR RADARES", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.write("Panel táctico de auditoría. Identifica vencimientos y ejecuta anulaciones de ingresos directo a la base maestra de SAP.")

    gc = inicializar_cliente_gspread()
    if not gc:
        st.error("🚨 Servidor desconectado. Revisa tus credenciales de Google Cloud.")
        return

    # URL Maestra que me proporcionaste
    URL_SHEET = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"

    with st.spinner("📡 Escaneando Bóveda de Ingresos en Drive..."):
        try:
            sh = gc.open_by_url(URL_SHEET)
            ws = sh.get_worksheet(0) # Se asume que la data está en la primera hoja (pestaña)
            datos_crudos = ws.get_all_values()
        except Exception as e:
            st.error(f"🚨 Error de acceso. Asegúrate de haber dado permisos de 'Editor' a la cuenta de servicio. Detalle: {e}")
            return

    if not datos_crudos or len(datos_crudos) < 2:
        st.warning("La base de datos parece estar vacía o no tiene encabezados válidos.")
        return

    encabezados = [str(x).strip().upper() for x in datos_crudos[0]]
    df = pd.DataFrame(datos_crudos[1:], columns=encabezados)
    
    # Inyectamos una columna invisible para saber exactamente qué fila de Excel editar (1-based index, saltando el header)
    df['FILA_EXCEL'] = range(2, len(df) + 2)

    # Validar que exista la columna ESTADO / OBSERVACIÓN
    COL_ESTADO = "ESTADO / OBSERVACIÓN"
    if COL_ESTADO not in df.columns:
        st.error(f"🚨 FALTA COLUMNA TÁCTICA: Debes crear una columna llamada exactamente **{COL_ESTADO}** en la fila 1 de tu Google Sheet.")
        return

    idx_col_estado = encabezados.index(COL_ESTADO) + 1 # Para gspread es base 1

    # --- 📊 1. PANEL DE RADARES (KPIs) ---
    st.markdown("### 📡 Radares de Vencimiento")
    
    hoy = datetime.now()
    limite_90_dias = hoy + timedelta(days=90)
    
    # Ubicar la columna F/V (Fecha Vencimiento)
    col_fv = next((c for c in df.columns if c in ["F/V", "FECHA VENCIMIENTO", "VENCIMIENTO"]), None)
    
    lotes_vencidos = 0
    lotes_riesgo = 0
    
    if col_fv:
        df['FECHA_VENC_DT'] = df[col_fv].apply(procesar_fecha_estricta)
        
        # Ignorar los ya anulados para la estadística
        df_activos = df[~df[COL_ESTADO].str.contains("ANULADO", na=False, case=False)]
        
        lotes_vencidos = df_activos[df_activos['FECHA_VENC_DT'] < hoy].shape[0]
        lotes_riesgo = df_activos[(df_activos['FECHA_VENC_DT'] >= hoy) & (df_activos['FECHA_VENC_DT'] <= limite_90_dias)].shape[0]

    k1, k2, k3 = st.columns(3)
    k1.markdown(f"""
    <div class='kpi-card kpi-verde'>
        <div class='kpi-titulo'>Total Ingresos Registrados</div>
        <p class='kpi-valor'>{len(df)}</p>
    </div>
    """, unsafe_allow_html=True)

    k2.markdown(f"""
    <div class='kpi-card kpi-rojo'>
        <div class='kpi-titulo'>🚨 Lotes Vencidos (Activos)</div>
        <p class='kpi-valor'>{lotes_vencidos}</p>
    </div>
    """, unsafe_allow_html=True)

    k3.markdown(f"""
    <div class='kpi-card kpi-amarillo'>
        <div class='kpi-titulo'>⚠️ Por Vencer (90 Días)</div>
        <p class='kpi-valor'>{lotes_riesgo}</p>
    </div>
    """, unsafe_allow_html=True)

    # --- 🛠️ 2. TABLA DE AUDITORÍA Y ANULACIONES ---
    st.markdown("---")
    st.markdown("### 🛠️ Matriz de Auditoría")
    st.caption("💡 Para anular un ingreso, haz doble clic en la columna **ESTADO / OBSERVACIÓN**, elige el motivo y presiona 'Guardar Anulaciones'.")

    # Limpiamos estados en blanco
    df[COL_ESTADO] = df[COL_ESTADO].replace(r'^\s*$', '✅ VIGENTE', regex=True).fillna('✅ VIGENTE')

    # Columnas a bloquear (todas excepto ESTADO)
    cols_disabled = [col for col in df.columns if col not in [COL_ESTADO, 'FILA_EXCEL', 'FECHA_VENC_DT']]
    
    # Opciones estandarizadas para el menú desplegable
    opciones_estado = [
        "✅ VIGENTE",
        "❌ ANULADO: ERROR EN PRECIOS",
        "❌ ANULADO: ERROR DE CANTIDAD",
        "❌ ANULADO: DEVOLUCIÓN A PROVEEDOR",
        "❌ ANULADO: ERROR EN LOTE/FECHAS",
        "❌ ANULADO: OTRO MOTIVO"
    ]

    # Preparamos vista
    columnas_vista = [c for c in df.columns if c not in ['FILA_EXCEL', 'FECHA_VENC_DT']]
    df_vista = df[columnas_vista].copy()

    df_editado = st.data_editor(
        df_vista,
        column_config={
            COL_ESTADO: st.column_config.SelectboxColumn(
                "ESTADO / OBSERVACIÓN",
                help="Selecciona el estado operativo de este ingreso.",
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

    # --- 💾 3. MOTOR DE SINCRONIZACIÓN (SALVAR EN DRIVE) ---
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 EJECUTAR Y SINCRONIZAR ANULACIONES EN DRIVE", type="primary"):
        cambios_detectados = False
        
        # Comparamos el dataframe original con el editado
        for i in range(len(df)):
            estado_original = str(df.iloc[i][COL_ESTADO]).strip()
            estado_nuevo = str(df_editado.iloc[i][COL_ESTADO]).strip()
            
            if estado_original != estado_nuevo:
                fila_excel = df.iloc[i]['FILA_EXCEL']
                with st.spinner(f"Inyectando cambio en Fila {fila_excel} de Google Sheets..."):
                    try:
                        # Actualizamos la celda exacta en Drive
                        ws.update_cell(fila_excel, idx_col_estado, estado_nuevo)
                        cambios_detectados = True
                    except Exception as e:
                        st.error(f"Error al actualizar fila {fila_excel}. Detalle: {e}")
        
        if cambios_detectados:
            st.success("✅ ¡Misión Cumplida! Bóveda de Drive actualizada exitosamente.")
            st.cache_data.clear() # Limpiar memoria para forzar recarga
            st.rerun()
        else:
            st.info("No se detectaron cambios en los estados. Nada que sincronizar.")

if __name__ == "__main__":
    ejecutar()
