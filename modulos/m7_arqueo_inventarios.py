import streamlit as st
import pandas as pd
from supabase import create_client, Client

# =================================================================
# 🔌 CONEXIÓN AL BÚNKER DE DATOS OMEGA PRO
# =================================================================
@st.cache_resource
def iniciar_conexion():
    url = st.secrets["SUPABASE_URL"].replace('"', '').replace("'", "").strip()
    key = st.secrets["SUPABASE_KEY"].replace('"', '').replace("'", "").strip()
    return create_client(url, key)

def ejecutar():
    st.markdown("<h1 style='color: #0d1b2a;'>⚖️ Módulo Gerencial: Arqueo de Inventarios</h1>", unsafe_allow_html=True)
    st.caption("Consolidación y cruce automático de inventarios: Planta física de aeródromos vs. SAP Libre Utilización.")

    try:
        supabase: Client = iniciar_conexion()
    except Exception:
        st.error("⚠️ Error de enlace con el centro de datos maestro.")
        return

    # Cargar base de inventario físico digitalizado en la app
    try:
        res_fisico = supabase.table("inventario_fisico").select("*").execute()
        datos_fisico = res_fisico.data if res_fisico.data else []
    except Exception as e:
        st.error(f"Error al extraer la base física: {e}")
        return

    # Pestañas operativas del Centro de Mando
    tab1, tab2, tab3 = st.tabs(["⚠️ Discrepancias Detectadas", "🔧 Conciliador de Carga", "📋 Inventario Completo"])

    # -----------------------------------------------------------------
    # PESTAÑA 2: EL CONCILIADOR (EL PUERTO DE CARGA SAP)
    # -----------------------------------------------------------------
    with tab2:
        st.markdown("### 📥 Inyección de Archivo Maestro de SAP")
        st.write("Cargue el reporte oficial de SAP (`.xlsx`) extraído del sistema central para actualizar la columna de Libre Utilización.")
        
        archivo_sap = st.file_uploader("Arrastre aquí el reporte EXPORT de SAP:", type=["xlsx"])
        
        if archivo_sap:
            try:
                # 🛡️ FIX 2: Configuración regional explícita para el surtido de miles (.) y decimales (,)
                df_sap = pd.read_excel(archivo_sap, thousands='.', decimal=',')
                
                st.success("✅ Documento de SAP analizado en memoria.")
                
                # Normalización de columnas críticas para evitar fallas por tildes o espacios
                df_sap.columns = df_sap.columns.str.strip()
                
                # Validar la existencia de la columna de la discordia
                col_sap_target = "Libre utilización"
                if col_sap_target not in df_sap.columns:
                    # Intento de rescate por si cambia de nombre
                    posibles_nombres = [c for c in df_sap.columns if "libre" in c.lower()]
                    if posibles_nombres:
                        col_sap_target = posibles_nombres[0]
                    else:
                        st.error("💥 Error Crítico: No se detectó la columna 'Libre utilización' en el archivo cargado.")
                        return

                if st.button("🚀 INYECTAR SALDOS DE LIBRE UTILIZACIÓN EN LA PLATAFORMA"):
                    progreso = st.progress(0)
                    actualizados = 0
                    
                    # Recorrer el archivo de SAP y actualizar la base de datos
                    for index, fila in df_sap.iterrows():
                        # 🛡️ FIX 3: Limpieza estricta de llaves de cruce para evitar saltos de formato
                        almacen_sap = str(fila.get("Almacén", "")).strip().upper()
                        lote_sap = str(fila.get("Lote", "")).strip().upper()
                        
                        # Manejo seguro del formato de números de libre utilización
                        saldo_libre_utilizacion = pd.to_numeric(fila.get(col_sap_target), errors='coerce')
                        if pd.isna(saldo_libre_utilizacion):
                            saldo_libre_utilizacion = 0.0

                        if almacen_sap and lote_sap and lote_sap != "NAN":
                            # Actualizar todas las filas físicas que coincidan con esa pista y lote
                            try:
                                supabase.table("inventario_fisico")\
                                    .update({"saldo_sap": saldo_libre_utilizacion})\
                                    .eq("pista", almacen_sap)\
                                    .eq("lote", lote_sap)\
                                    .execute()
                                actualizados += 1
                            except Exception:
                                pass
                                
                    st.success(f"🎉 ¡Sincronización Terminada! Se procesaron y mapearon {actualizados} lotes de SAP con éxito.")
                    st.rerun()

            except Exception as e:
                st.error(f"Falla al decodificar las matrices de SAP: {e}")

    # -----------------------------------------------------------------
    # PROCESAMIENTO Y ARMADO DE LA MATRIZ GENERAL DE ARQUEO
    # -----------------------------------------------------------------
    if datos_fisico:
        df_master = pd.DataFrame(datos_fisico)
        
        # Saneamiento y casteo preventivo de variables financieras y de volumen
        df_master["saldo_sap"] = pd.to_numeric(df_master.get("saldo_sap", 0.0), errors="coerce").fillna(0.0)
        df_master["saldo_fisico"] = pd.to_numeric(df_master.get("saldo_fisico", 0.0), errors="coerce").fillna(0.0)
        
        # Cálculo en tiempo real de la brecha logística
        df_master["Diferencia"] = df_master["saldo_fisico"] - df_master["saldo_sap"]
        
        # Marcación de estados semafóricos
        def determinar_estado(dif):
            return "✅ OK" if abs(dif) < 0.001 else "❌ DISCREPANCIA"
            
        df_master["Estado"] = df_master["Diferencia"].apply(determinar_estado)
        
        # Filtros de visualización estratégica
        df_discrepancias = df_master[df_master["Estado"] == "❌ DISCREPANCIA"].copy()
    else:
        df_master = pd.DataFrame()
        df_discrepancias = pd.DataFrame()

    # -----------------------------------------------------------------
    # PESTAÑA 1: TABLERO DE CONTROL DE DISCREPANCIAS
    # -----------------------------------------------------------------
    with tab1:
        # KPI Cards Superiores Avanzados
        total_items = len(df_master)
        items_ok = len(df_master[df_master["Estado"] == "✅ OK"])
        total_alarmas = len(df_discrepancias)
        balance_fisico = df_master["saldo_fisico"].sum() if not df_master.empty else 0.0

        # Estilización premium de tarjetas de balance
        st.markdown(f"""
        <div style="background-color: #0d1b2a; padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 2px solid #d4af37;">
            <table style="width: 100%; border: none; text-align: center; color: white;">
                <tr>
                    <td><strong>LOTES ARQUEADOS</strong><br><span style="font-size: 20px; color: #d4af37;">⚖️ {total_items} Ítems</span></td>
                    <td><strong>CUADRADOS CON SAP</strong><br><span style="font-size: 20px; color: #2b9348;">🟢 {items_ok} OK</span></td>
                    <td><strong>DESFASES CRÍTICOS</strong><br><span style="font-size: 20px; color: #e63946;">⚠️ {total_alarmas} Alarmas</span></td>
                    <td><strong>BALANCE NETO FÍSICO</strong><br><span style="font-size: 20px; color: #219ebc;">📥 {balance_fisico:,.2f} L/Kg</span></td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🔍 Listado de Desfases Logísticos Activos")
        if not df_discrepancias.empty:
            # Estructurar la visualización exacta que requieres en pantalla
            df_view = df_discrepancias[['pista', 'producto', 'lote', 'saldo_sap', 'saldo_fisico', 'Diferencia', 'Estado']].copy()
            df_view.columns = ['PISTA', 'PRODUCTO', 'LOTE', 'SALDO SAP', 'SALDO FÍSICO', 'DIFERENCIA', 'ESTADO']
            
            # Formatear números para visualización limpia de 3 decimales
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
            df_total_view = df_master[['pista', 'producto', 'lote', 'saldo_sap', 'saldo_fisico', 'Diferencia', 'Estado']].copy()
            df_total_view.columns = ['PISTA', 'PRODUCTO', 'LOTE', 'SALDO SAP', 'SALDO FÍSICO', 'DIFERENCIA', 'ESTADO']
            st.dataframe(df_total_view.sort_values(by=["PISTA", "PRODUCTO"]), use_container_width=True, hide_index=True)
        else:
            st.info("No hay registros en la bitácora de inventarios.")

if __name__ == "__main__":
    ejecutar()
