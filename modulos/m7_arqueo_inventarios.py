import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 CONFIGURACIÓN DE TU CENTRAL DE GOOGLE SHEETS
# =================================================================
# ⚠️ REVISA ESTOS DOS NOMBRES EXACTAMENTE COMO SE LLAMAN EN TU DRIVE
NOMBRE_DEL_DRIVE = "Génesis Omega Pro" 
NOMBRE_DE_LA_HOJA = "inventario_fisico"

@st.cache_resource
def iniciar_conexion_google():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ Error en las credenciales de Google Sheets: {e}")
        return None

def ejecutar(quitar_tildes=None, purificar_lote=None):
    st.markdown("<h1 style='color: #0d1b2a;'>⚖️ Módulo Gerencial: Arqueo de Inventarios</h1>", unsafe_allow_html=True)
    st.caption("Consolidación y cruce automático de inventarios: Planta física de aeródromos vs. SAP Libre Utilización.")

    cliente_google = iniciar_conexion_google()
    if not cliente_google:
        st.error("❌ No se pudo establecer conexión con el entorno de Google.")
        return

    # =================================================================
    # 🔍 AUDITORÍA PASO A PASO DE TU ARCHIVO DE GOOGLE DRIVE
    # =================================================================
    
    # PASO 1: Intentar abrir el archivo maestro
    try:
        libro = cliente_google.open(NOMBRE_DEL_DRIVE)
    except Exception as e:
        st.error(f"❌ PASO 1 FALLIDO: No se encontró el archivo '{NOMBRE_DEL_DRIVE}' en Google Drive.")
        st.info("💡 **Solución:** Asegúrese de que el archivo en Drive se llame exactamente así (ojo con las tildes o espacios) y que esté compartido con el correo de su cuenta de servicio (`client_email` del JSON).")
        return

    # PASO 2: Intentar abrir la pestaña específica
    try:
        worksheet = libro.worksheet(NOMBRE_DE_LA_HOJA)
    except Exception as e:
        st.error(f"❌ PASO 2 FALLIDO: El archivo abrió bien, pero no existe la pestaña '{NOMBRE_DE_LA_HOJA}'.")
        st.info("💡 **Solución:** Revise el nombre de la pestaña abajo en su Google Sheet. Debe ser todo en minúsculas y con guión bajo.")
        return

    # PASO 3: Intentar extraer las filas de la tabla
    try:
        datos_gspread = worksheet.get_all_records()
    except Exception as e:
        st.error(f"❌ PASO 3 FALLIDO: Se conectó al archivo y a la pestaña, pero falló la lectura de celdas.")
        st.warning(f"Detalle técnico del error: {e}")
        st.info("💡 **Solución:** Esto ocurre si la hoja de cálculo está completamente vacía (sin columnas de encabezados) o si el archivo subido a Drive es un `.xlsx` puro y no ha sido convertido al formato nativo de Google Sheets.")
        return

    # Pestañas operativas del Centro de Mando
    tab1, tab2, tab3 = st.tabs(["⚠️ Discrepancias Detectadas", "🔧 Conciliador de Carga", "📋 Inventario Completo"])

    # -----------------------------------------------------------------
    # PESTAÑA 2: EL CONCILIADOR (EL PUERTO DE CARGA SAP)
    # -----------------------------------------------------------------
    with tab2:
        st.markdown("### 📥 Inyección de Archivo Maestro de SAP")
        st.write("Cargue el reporte oficial de SAP (`.xlsx`) para actualizar la columna de Libre Utilización.")
        
        archivo_sap = st.file_uploader("Arrastre aquí el reporte EXPORT de SAP:", type=["xlsx"])
        
        if archivo_sap:
            try:
                df_sap = pd.read_excel(archivo_sap, thousands='.', decimal=',')
                st.success("✅ Documento de SAP analizado en memoria.")
                
                df_sap.columns = df_sap.columns.str.strip()
                col_sap_target = "Libre utilización"
                
                if col_sap_target not in df_sap.columns:
                    posibles_nombres = [c for c in df_sap.columns if "libre" in c.lower()]
                    if posibles_nombres:
                        col_sap_target = posibles_nombres[0]
                    else:
                        st.error("💥 Error Crítico: No se detectó la columna 'Libre utilización' en el archivo.")
                        return

                if st.button("🚀 INYECTAR SALDOS DE LIBRE UTILIZACIÓN EN GOOGLE SHEETS"):
                    df_fisico_db = pd.DataFrame(datos_gspread)
                    df_fisico_db.columns = df_fisico_db.columns.str.strip()
                    
                    if "saldo_sap" not in df_fisico_db.columns:
                        df_fisico_db["saldo_sap"] = 0.0

                    actualizados = 0
                    
                    with st.spinner("Cruzando matrices y actualizando celdas..."):
                        for index, fila in df_sap.iterrows():
                            almacen_raw = str(fila.get("Almacén", "")).strip().upper()
                            almacen_sap = quitar_tildes(almacen_raw) if quitar_tildes else almacen_raw
                            
                            lote_raw = str(fila.get("Lote", "")).strip().upper()
                            lote_sap = purificar_lote(lote_raw) if purificar_lote else lote_raw
                            
                            saldo_libre_utilizacion = pd.to_numeric(fila.get(col_sap_target), errors='coerce')
                            if pd.isna(saldo_libre_utilizacion):
                                saldo_libre_utilizacion = 0.0

                            mascara = (df_fisico_db["pista"].astype(str).str.strip().str.upper() == almacen_sap) & \
                                      (df_fisico_db["lote"].astype(str).str.strip().str.upper() == lote_sap)
                            
                            if mascara.any():
                                df_fisico_db.loc[mascara, "saldo_sap"] = saldo_libre_utilizacion
                                actualizados += 1

                        worksheet.clear()
                        df_subida = df_fisico_db.fillna("")
                        worksheet.update([df_subida.columns.values.tolist()] + df_subida.values.tolist())
                        
                    st.success(f"🎉 ¡Sincronización Terminada! Se mapearon {actualizados} registros en tu Google Sheet.")
                    st.rerun()

            except Exception as e:
                st.error(f"Falla al procesar las celdas: {e}")

    # -----------------------------------------------------------------
    # PROCESAMIENTO GENERAL DE MATRICES
    # -----------------------------------------------------------------
    if datos_gspread:
        df_master = pd.DataFrame(datos_gspread)
        df_master.columns = df_master.columns.str.strip()
        
        df_master["saldo_sap"] = pd.to_numeric(df_master.get("saldo_sap", 0.0), errors="coerce").fillna(0.0)
        df_master["saldo_fisico"] = pd.to_numeric(df_master.get("saldo_fisico", 0.0), errors="coerce").fillna(0.0)
        df_master["Diferencia"] = df_master["saldo_fisico"] - df_master["saldo_sap"]
        
        def determinar_estado(dif):
            return "✅ OK" if abs(dif) < 0.001 else "❌ DISCREPANCIA"
            
        df_master["Estado"] = df_master["Diferencia"].apply(determinar_estado)
        df_discrepancias = df_master[df_master["Estado"] == "❌ DISCREPANCIA"].copy()
    else:
        df_master = pd.DataFrame()
        df_discrepancias = pd.DataFrame()

    # -----------------------------------------------------------------
    # PESTAÑA 1: TABLERO DE CONTROL DE DISCREPANCIAS
    # -----------------------------------------------------------------
    with tab1:
        total_items = len(df_master)
        items_ok = len(df_master[df_master["Estado"] == "✅ OK"]) if not df_master.empty else 0
        total_alarmas = len(df_discrepancias)
        balance_fisico = df_master["saldo_fisico"].sum() if not df_master.empty else 0.0

        st.markdown(f"""
        <div style="background-color: #0d1b2a; padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 2px solid #d4af37;">
            <table style="width: 100%; border: none; text-align: center; color: white;">
                <tr>
                    <td><strong>LOTES ARQUEADOS</strong><br><span style="font-size: 20px; color: #d4af37;">⚖️ {total_items} Ítems</span></td>
                    <td><strong>CUADRADOS CON SAP</strong><br><span style="font-size: 20px; color: #2b9348;">🟢 {items_ok} OK</span></td>
                    <td><strong>DESFACES CRÍTICOS</strong><br><span style="font-size: 20px; color: #e63946;">⚠️ {total_alarmas} Alarmas</span></td>
                    <td><strong>BALANCE NETO FÍSICO</strong><br><span style="font-size: 20px; color: #219ebc;">📥 {balance_fisico:,.2f} L/Kg</span></td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🔍 Listado de Desfases Logísticos Activos")
        if not df_discrepancias.empty:
            columnas_existentes = df_discrepancias.columns
            pista_col = "pista" if "pista" in columnas_existentes else columnas_existentes[0]
            prod_col = "producto" if "producto" in columnas_existentes else columnas_existentes[1]
            lote_col = "lote" if "lote" in columnas_existentes else columnas_existentes[2]

            df_view = df_discrepancias[[pista_col, prod_col, lote_col, 'saldo_sap', 'saldo_fisico', 'Diferencia', 'Estado']].copy()
            df_view.columns = ['PISTA', 'PRODUCTO', 'LOTE', 'SALDO SAP', 'SALDO FÍSICO', 'DIFERENCIA', 'ESTADO']
            
            for col in ['SALDO SAP', 'SALDO FÍSICO', 'DIFERENCIA']:
                df_view[col] = df_view[col].map(lambda x: f"{x:.3f}")
                
            st.dataframe(df_view, use_container_width=True, hide_index=True)
        else:
            st.success("🎯 Sistema en Balance Absoluto: No se registran discrepancias con SAP en ninguna pista.")

    # -----------------------------------------------------------------
    # PESTAÑA 3: INVENTARIO CENTRALIZADO TOTAL
    # -----------------------------------------------------------------
    with tab3:
        st.markdown("### 📋 Historial y Bitácora Completa de Existencias")
        if not df_master.empty:
            columnas_existentes = df_master.columns
            pista_col = "pista" if "pista" in columnas_existentes else columnas_existentes[0]
            prod_col = "producto" if "producto" in columnas_existentes else columnas_existentes[1]
            lote_col = "lote" if "lote" in columnas_existentes else columnas_existentes[2]

            df_total_view = df_master[[pista_col, prod_col, lote_col, 'saldo_sap', 'saldo_fisico', 'Diferencia', 'Estado']].copy()
            df_total_view.columns = ['PISTA', 'PRODUCTO', 'LOTE', 'SALDO SAP', 'SALDO FÍSICO', 'DIFERENCIA', 'ESTADO']
            st.dataframe(df_total_view.sort_values(by=["PISTA", "PRODUCTO"]), use_container_width=True, hide_index=True)
        else:
            st.info("No hay registros en la bitácora de inventarios.")

if __name__ == "__main__":
    ejecutar()
