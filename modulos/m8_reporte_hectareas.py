import streamlit as st
import pandas as pd

# =================================================================
# 🎣 CÓDIGO CEBO PARA DETECTAR LA MUERTE SÚBITA
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    st.markdown("<h1>Radar de Hectáreas - MODO RASTREO</h1>", unsafe_allow_html=True)
    
    st.warning("🕵️‍♂️ INICIANDO OPERACIÓN DE RASTREO...")
    
    # --- CEBO 1 ---
    st.info("🎣 CEBO 1: El enrutador funcionó y entramos al Módulo 8 correctamente.")

    if supabase_client is None:
        st.error("🚨 CEBO FALLIDO: El cliente de Supabase llegó vacío desde app.py.")
        return

    try:
        # --- CEBO 2 ---
        st.info("🎣 CEBO 2: Tocando la puerta de Supabase para pedir solo 5 filas...")
        respuesta = supabase_client.table("TABLA_1").select("*").limit(5).execute()
        raw_data = respuesta.data
        
        # --- CEBO 3 ---
        st.info(f"🎣 CEBO 3: Supabase respondió. Se recibieron {len(raw_data)} filas.")
        
        if not raw_data:
            st.warning("⚠️ La TABLA_1 está vacía. Fin de la prueba.")
            return
            
        # --- CEBO 4 ---
        st.info("🎣 CEBO 4: Convirtiendo datos a Pandas DataFrame...")
        df_raw = pd.DataFrame(raw_data)
        
        # --- CEBO 5 ---
        st.info("🎣 CEBO 5: Forzando a texto puro para evitar colapsos gráficos y mostrando en pantalla:")
        
        # Burlamos a PyArrow forzando todo a texto
        st.write(df_raw.astype(str))
        
        # --- CEBO 6 ---
        st.success("🎉 CEBO 6: ¡Misión Cumplida! Si ves este mensaje verde, la conexión, la memoria y la tabla funcionan perfectamente.")

    except Exception as e:
        st.error(f"🚨 EL CEBO ATRAPÓ UN ERROR FATAL: {e}")

if __name__ == "__main__":
    pass
