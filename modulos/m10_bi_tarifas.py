import streamlit as st
import pandas as pd
import gspread

st.set_page_config(layout="wide")

def cargar_datos_google(url, hoja):
    if "gcp_credentials" in st.secrets:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
    else:
        gc = gspread.service_account(filename="credenciales.json")
    sh = gc.open_by_url(url)
    ws = sh.worksheet(hoja)
    datos = ws.get_all_values()
    if not datos: return pd.DataFrame()
    df = pd.DataFrame(datos[1:], columns=datos[0])
    return df

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

def main():
    st.title("🪤 TRAMPA DE DEPURACIÓN (RAYOS X)")
    st.info("Descargando datos crudos directamente desde Drive...")

    url_actual = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    df = cargar_datos_google(url_actual, "TABLA 1")

    if df.empty:
        st.error("No hay datos.")
        return

    # Limpiar encabezados básico
    df.columns = [str(c).upper().strip() for c in df.columns]

    # Encontrar las columnas dinámicamente
    col_w, col_f, col_finca = None, None, None
    for c in df.columns:
        if 'FACTURAR' in c and ('PRODUCTOR' in c or 'CICLO' in c): col_w = c
        if 'FUMIG' in c and 'AREA' in c: col_f = c
        if 'FINCA' in c or 'PROPIEDAD' in c: col_finca = c

    if not col_w or not col_f:
        st.error(f"No encontré las columnas. W: {col_w}, F: {col_f}")
        st.write("Columnas disponibles:", df.columns.tolist())
        return

    # Construir la tabla trampa
    df_trampa = pd.DataFrame()
    df_trampa['FINCA'] = df[col_finca] if col_finca else "N/A"
    
    # Textos Crudos (COMO VIENEN DE GOOGLE)
    df_trampa['AREA_CRUDA (TEXTO)'] = df[col_f]
    df_trampa['COSTO_CRUDO_W (TEXTO)'] = df[col_w]
    
    # Números procesados por Python
    df_trampa['AREA_PYTHON (NUM)'] = df[col_f].apply(limpiar_area_trampa)
    df_trampa['COSTO_PYTHON_W (NUM)'] = df[col_w].apply(convertir_pesos_trampa)

    # Filtrar para no ver filas vacías y enfocarnos en donde hay dinero
    df_trampa = df_trampa[df_trampa['COSTO_PYTHON_W (NUM)'] > 0]

    st.success("✅ Datos capturados. Revise la comparación cara a cara:")
    st.dataframe(df_trampa.head(100), use_container_width=True)

if __name__ == "__main__":
    main()
