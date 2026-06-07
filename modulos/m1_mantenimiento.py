import streamlit as st
import pandas as pd
import gspread
import io
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# ⚡ MOTORES DE CONEXIÓN PROPIO (ANTENA SATELITAL DE ALTA VELOCIDAD)
# =================================================================

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread_propio():
    """ Centraliza la autenticación con Google Cloud usando el llavero unificado """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # 🌟 COMPATIBILIDAD ABSOLUTA: Usamos el secreto que ya dejamos funcionando
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
        return None
    except:
        return None

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE LA PLANTILLA MAESTRA DE SAP
# =================================================================

def ejecutar(quitar_tildes=None, purificar_lote=None):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>INTELEGENCIA DE PRECIOS SAP</h1>", unsafe_allow_html=True)
    
    # Prueba de conexión satelital nativa
    gc = inicializar_cliente_gspread_propio()
    if gc is None:
        st.error("🚨 No se pudo establecer conexión con Google Cloud. Verifique sus credenciales.")
        return

    st.markdown("### 📥 1. Suba la Sábana Cruda de SAP")
    archivo_sap = st.file_uploader("Suba el reporte oficial extraído del sistema central (.xlsx):", type=['xlsx', 'csv'], key="uploader_m1")

    if archivo_sap:
        st.success("✅ Archivo maestro cargado en la memoria temporal de la app.")
        
        if st.button("🚀 PASO A: PURIFICAR Y CARGAR A PLANTILLA", type="primary", use_container_width=True):
            try:
                with st.spinner("Procesando matriz de precios y depurando lotes..."):
                    # Leer el archivo de SAP de forma regional
                    if archivo_sap.name.lower().endswith('.xlsx') or archivo_sap.name.lower().endswith('.xls'):
                        df_sap = pd.read_excel(archivo_sap)
                    else:
                        df_sap = pd.read_csv(archivo_sap, sep=None, engine='python', encoding='utf-8')

                    # Limpiar encabezados de espacios o tildes
                    df_sap.columns = [quitar_tildes(str(c)).strip() if quitar_tildes else str(c).strip() for c in df_sap.columns]
                    
                    # 🔍 Conexión directa al Drive Destino para el Mantenimiento
                    # OJO: Modifica el nombre o URL si la base de precios es otra hoja
                    url_base_precios = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    sh = gc.open_by_url(url_base_precios)
                    
                    # Buscamos la pestaña correspondiente (ej: "Plantilla SAP" o la primera por defecto)
                    try:
                        ws_plantilla = sh.worksheet("Plantilla SAP")
                    except:
                        ws_plantilla = sh.get_worksheet(0)

                    # --- ALGORITMO DE PURIFICACIÓN DE LA SÁBANA ---
                    df_sap_clean = df_sap.copy()
                    
                    # Si tu app principal tiene filtros de lote o almacén, los aplicamos aquí
                    if purificar_lote and "Lote" in df_sap_clean.columns:
                        df_sap_clean["Lote_Limpio"] = df_sap_clean["Lote"].apply(purificar_lote)

                    # Reemplazar NaN por texto vacío antes de subir a la nube
                    df_subida = df_sap_clean.fillna("")
                    
                    # Volcado masivo y limpio a Google Sheets
                    ws_plantilla.clear()
                    ws_plantilla.update([df_subida.columns.values.tolist()] + df_subida.values.tolist())
                    
                    st.balloons()
                    st.success(f"🎯 ¡OPERACIÓN EXITOSA! La pestaña '{ws_plantilla.title}' en Google Drive ha sido actualizada con {len(df_subida)} registros purificados.")
                    
            except Exception as e:
                st.error(f"🚨 Error durante la inyección de la plantilla: {e}")

if __name__ == "__main__":
    pass
