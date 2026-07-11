import streamlit as st

# =================================================================
# 🕵️‍♂️ ESCÁNER FORENSE (SIN PANDAS / SIN PYARROW)
# =================================================================
def ejecutar(supabase_client, descargar_matriz_rapida=None, extraer_numero_ext=None, procesar_fecha_pesada_ext=None, HAS_MATPLOTLIB=True):
    st.markdown("<h1>Radar - MODO FORENSE</h1>", unsafe_allow_html=True)
    
    if supabase_client is None:
        st.error("🚨 Sin conexión a Supabase.")
        return

    try:
        st.info("🔍 PASO 1: Pidiendo 1 sola fila a Supabase...")
        respuesta = supabase_client.table("TABLA_1").select("*").limit(1).execute()
        datos = respuesta.data
        
        st.info("🔍 PASO 2: Fila recibida. Desarmando columna por columna (Sin Pandas)...")
        
        if not datos:
            st.warning("La tabla está vacía.")
            return
            
        fila = datos[0]
        
        st.markdown("### ☢️ Radiografía de la Fila 1:")
        # Imprimimos cada columna como texto puro para evitar colapsos
        for nombre_columna, valor_celda in fila.items():
            tipo_dato = type(valor_celda).__name__
            st.code(f"COLUMNA: {nombre_columna} | VALOR: {valor_celda} | TIPO: {tipo_dato}", language="text")
            
        st.success("✅ PASO 3: ¡SUPERVIVENCIA CONFIRMADA! Si ves este mensaje verde, el problema es 100% una incompatibilidad de PyArrow.")

    except Exception as e:
        st.error(f"🚨 EL ESCÁNER DETECTÓ UN ERROR: {e}")

if __name__ == "__main__":
    pass
