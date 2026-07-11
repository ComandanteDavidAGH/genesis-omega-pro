import streamlit as st

def ejecutar(supabase_client, *args):
    st.error("🛑 MODO DE DIAGNÓSTICO PROFUNDO")
    
    st.info("PUNTO 1: Código iniciado. Conectando a Supabase...")
    
    if supabase_client is None:
        st.error("Fallo de conexión.")
        return

    try:
        # Pedimos 50 filas
        respuesta = supabase_client.table("TABLA_1").select("*").limit(50).execute()
        data = respuesta.data
        
        st.success("PUNTO 2: Datos recibidos desde Supabase sin colapsar la red.")
        
        if not data:
            st.warning("La tabla está vacía.")
            return

        st.warning("PUNTO 3: Imprimiendo datos crudos en la pantalla...")
        
        # Imprimimos los datos en formato crudo de texto (sin tablas ni Pandas)
        st.json(data)
        
        st.success("PUNTO 4: ¡SI VES ESTO, LA APP NO MURIÓ Y EL PROBLEMA ERA PANDAS!")
            
    except Exception as e:
        st.error(f"Error detectado: {e}")

if __name__ == "__main__":
    pass
