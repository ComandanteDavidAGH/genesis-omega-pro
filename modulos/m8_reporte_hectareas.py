import streamlit as st
import pandas as pd

def ejecutar(supabase_client, *args):
    st.title("📡 Radar - Modo Supervivencia")
    
    if supabase_client is None:
        st.error("El cliente de Supabase no se recibió.")
        return

    st.info("Intentando descargar 50 filas de la TABLA_1...")

    try:
        # Llamada directa sin procesamiento pesado
        respuesta = supabase_client.table("TABLA_1").select("*").limit(50).execute()
        data = respuesta.data
        
        if not data:
            st.warning("La tabla está vacía.")
        else:
            # Convertir a DataFrame forzando todo a string (evita errores de tipo en PyArrow)
            df = pd.DataFrame(data)
            st.success(f"¡Éxito! Se descargaron {len(df)} filas.")
            st.dataframe(df.astype(str))
            
    except Exception as e:
        st.error(f"Error crítico durante la carga: {e}")

if __name__ == "__main__":
    pass
