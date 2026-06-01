import streamlit as st
import pandas as pd

def ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero):
    st.title("🪤 TRAMPA 2.0 (RADAR DE ENCABEZADOS)")
    st.info("Escaneando las primeras 15 filas del documento crudo...")

    try:
        url_actual = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
        datos = descargar_matriz_rapida(url_actual, "TABLA 1")
        
        if not datos:
            st.error("El documento está totalmente vacío.")
            return

        st.markdown("### 🔍 RADIOGRAFÍA DE LAS PRIMERAS 15 FILAS:")
        # Mostramos literalmente qué hay en cada fila al inicio del Excel
        for i, fila in enumerate(datos[:15]):
            st.text(f"Fila Excel {i+1} (Índice Python {i}): {fila}")

        st.markdown("---")
        # El Sabueso: Busca en qué fila exacta está la palabra FECHA o FINCA
        fila_titulos = -1
        for i, fila in enumerate(datos[:20]):
            fila_upper = [str(celda).upper().strip() for celda in fila]
            if "FECHA" in fila_upper or "FINCA" in fila_upper or "PROPIEDAD" in fila_upper:
                fila_titulos = i
                break
        
        if fila_titulos != -1:
            st.success(f"✅ **¡MISTERIO RESUELTO!** Los títulos reales están en la **Fila Excel {fila_titulos + 1}** (Índice Python {fila_titulos}).")
            st.write("Estos son los encabezados que encontró:")
            st.write(datos[fila_titulos])
        else:
            st.error("🚨 ALERTA: No encontré la palabra 'FECHA' ni 'FINCA' en las primeras 20 filas. ¿Seguro que estamos en la pestaña correcta?")

    except Exception as e:
        st.error(f"Error en el radar: {e}")
