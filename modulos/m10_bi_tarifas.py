import streamlit as st
import pandas as pd
import gspread

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
    st.info("Alineando la mira a la Fila 5 del Excel...")

    try:
        url_actual = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
        datos = descargar_matriz_rapida(url_actual, "TABLA 1")
        
        if not datos or len(datos) < 6:
            st.error("No hay suficientes datos en la Tabla 1.")
            return
            
        # 🎯 EL AJUSTE: Leemos desde la Fila 5 (índice 4 en Python)
        df = pd.DataFrame(datos[5:], columns=datos[4])
        df.columns = [str(c).upper().strip() for c in df.columns]

        col_w, col_f, col_finca, col_fecha = None, None, None, None
        for c in df.columns:
            # Ignoramos la V por seguridad
            if 'FINCA' in c and 'COSTO' in c: continue
            
            if 'FACTURAR' in c and ('PRODUCTOR' in c or 'CICLO' in c): col_w = c
            if 'FUMIG' in c and 'AREA' in c: col_f = c
            if 'FINCA' in c or 'PROPIEDAD' in c: col_finca = c
            if 'FECHA' == c: col_fecha = c

        if not col_w or not col_f:
            st.error(f"Sigo ciego. Encontré esto: W={col_w}, F={col_f}")
            st.write("Estos son los encabezados que veo en la Fila 5:", df.columns.tolist())
            return

        df_trampa = pd.DataFrame()
        df_trampa['FECHA'] = df[col_fecha] if col_fecha else "N/A"
        df_trampa['FINCA'] = df[col_finca] if col_finca else "N/A"
        
        df_trampa['AREA_CRUDA (TEXTO)'] = df[col_f]
        df_trampa['COSTO_CRUDO_W (TEXTO)'] = df[col_w]
        
        df_trampa['AREA_PYTHON (NUM)'] = df[col_f].apply(limpiar_area_trampa)
        df_trampa['COSTO_PYTHON_W (NUM)'] = df[col_w].apply(convertir_pesos_trampa)

        df_trampa = df_trampa[df_trampa['COSTO_PYTHON_W (NUM)'] > 0]

        st.success("✅ ¡Encabezados encontrados! Revise la comparación cara a cara:")
        st.dataframe(df_trampa.head(100), use_container_width=True)

    except Exception as e:
        st.error(f"Error en la trampa: {e}")
