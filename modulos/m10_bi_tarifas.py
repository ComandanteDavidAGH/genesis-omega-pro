import streamlit as st
import pandas as pd
import gspread

# Traductores bajo sospecha (los que vamos a vigilar)
def convertir_pesos_trampa(val):
    try:
        v = str(val).strip()
        if not v: return 0.0
        v = "".join([c for c in v if c.isdigit() or c in ['.', ',']])
        if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
        elif ',' in v:
            partes = v.split(',')
            if len(partes[-1]) == 3: v = v.replace(',', '')
            else: v = v.replace(',', '.')
        elif '.' in v:
            partes = v.split('.')
            if len(partes[-1]) == 3: v = v.replace('.', '')
        return float(v)
    except: return 0.0

def limpiar_area_trampa(val):
    try:
        v = str(val).strip()
        if not v: return 0.0
        v = "".join([c for c in v if c.isdigit() or c in ['.', ',']])
        if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
        elif ',' in v: v = v.replace(',', '.')
        return float(v)
    except: return 0.0

def ejecutar(descargar_matriz_rapida, procesar_fecha_pesada, extraer_numero):
    st.title("🪤 TRAMPA DE DEPURACIÓN (RAYOS X)")
    st.info("Descargando datos crudos directamente desde Drive...")

    try:
        url_actual = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
        # Usamos su función global para descargar
        datos = descargar_matriz_rapida(url_actual, "TABLA 1")
        if not datos or len(datos) < 2:
            st.error("No hay datos en la Tabla 1.")
            return
            
        df = pd.DataFrame(datos[1:], columns=datos[0])

        # Limpiar encabezados básico
        df.columns = [str(c).upper().strip() for c in df.columns]

        # Encontrar las columnas dinámicamente
        col_w, col_f, col_finca, col_fecha = None, None, None, None
        for c in df.columns:
            if 'FACTURAR' in c and ('PRODUCTOR' in c or 'CICLO' in c): col_w = c
            if 'FUMIG' in c and 'AREA' in c: col_f = c
            if 'FINCA' in c or 'PROPIEDAD' in c: col_finca = c
            if 'FECHA' == c: col_fecha = c

        if not col_w or not col_f:
            st.error(f"No encontré las columnas. W: {col_w}, F: {col_f}")
            st.write("Columnas disponibles:", df.columns.tolist())
            return

        # Construir la tabla trampa
        df_trampa = pd.DataFrame()
        df_trampa['FECHA'] = df[col_fecha] if col_fecha else "N/A"
        df_trampa['FINCA'] = df[col_finca] if col_finca else "N/A"
        
        # Textos Crudos (COMO VIENEN DE GOOGLE)
        df_trampa['AREA_CRUDA (TEXTO)'] = df[col_f]
        df_trampa['COSTO_CRUDO_W (TEXTO)'] = df[col_w]
        
        # Números procesados por Python
        df_trampa['AREA_PYTHON (NUM)'] = df[col_f].apply(limpiar_area_trampa)
        df_trampa['COSTO_PYTHON_W (NUM)'] = df[col_w].apply(convertir_pesos_trampa)

        # Filtrar para enfocarnos en los datos de 2026 o los que tienen dinero
        df_trampa = df_trampa[df_trampa['COSTO_PYTHON_W (NUM)'] > 0]

        st.success("✅ Datos capturados. Revise la comparación cara a cara:")
        st.dataframe(df_trampa, use_container_width=True)

    except Exception as e:
        st.error(f"Error en la trampa: {e}")
