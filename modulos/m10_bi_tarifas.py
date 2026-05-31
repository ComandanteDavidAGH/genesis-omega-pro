import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from datetime import datetime

# --- Funciones de limpieza ---
def limpiar_encabezados(df):
    df.columns = [
        str(col).upper()
        .replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
        .strip()
        for col in df.columns
    ]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    if "" in df.columns: df = df.drop(columns=[""])
    return df

def convertir_pesos(val):
    try:
        v = str(val)
        v_limpio = "".join([c for c in v if c.isdigit() or c in ['.', ',']])
        v_limpio = v_limpio.rstrip('.,')
        if v_limpio == '': return 0.0
        if ',' in v_limpio and '.' not in v_limpio: v_limpio = v_limpio.replace(',', '.')
        elif '.' in v_limpio and ',' in v_limpio: v_limpio = v_limpio.replace('.', '').replace(',', '.')
        return float(v_limpio)
    except:
        return 0.0

def limpiar_area(val):
    try:
        v = str(val).upper().replace(',', '.')
        v = "".join([c for c in v if c.isdigit() or c == '.'])
        return float(v) if v != '' else 0.0
    except:
        return 0.0

# --- Conexión a Google Sheets ---
def cargar_datos_google(url, hoja):
    # Autenticación con credenciales
    if "gcp_credentials" in st.secrets:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
    else:
        gc = gspread.service_account(filename="credenciales.json")

    sh = gc.open_by_url(url)
    ws = sh.worksheet(hoja)
    datos = ws.get_all_values()
    if not datos: return pd.DataFrame()
    df = pd.DataFrame(datos[1:], columns=datos[0])
    return limpiar_encabezados(df)

# --- Aplicación principal ---
def main():
    st.title("📊 Centro de Inteligencia Estratégica BI")
    st.info("Conectado directamente a Google Sheets")

    # URLs de tus hojas
    url_actual = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    url_historico = "https://docs.google.com/spreadsheets/d/16OZdiWwW7nLHyZBEnhiKlDTDttR7Tjhn37O9zm6wJOk/edit"

    df_actual = cargar_datos_google(url_actual, "TABLA 1")
    df_historico = cargar_datos_google(url_historico, "Datos")

    # Unir bases
    super_base = pd.concat([df_actual, df_historico], ignore_index=True)

    # Conversión de columnas clave
    if 'COSTO' in super_base.columns:
        super_base['COSTO_NUM'] = super_base['COSTO'].apply(convertir_pesos)
    if 'AREA' in super_base.columns:
        super_base['AREA_NUM'] = super_base['AREA'].apply(limpiar_area)
    if 'FECHA' in super_base.columns:
        super_base['FECHA'] = pd.to_datetime(super_base['FECHA'], errors='coerce')

    st.subheader("Vista previa de los datos")
    st.dataframe(super_base.head())

    # Métricas
    if 'COSTO_NUM' in super_base.columns:
        costo_promedio = super_base['COSTO_NUM'].mean()
        st.metric("Costo Promedio por Ha", f"$ {costo_promedio:,.0f}")

    if 'AREA_NUM' in super_base.columns:
        area_total = super_base['AREA_NUM'].sum()
        st.metric("Área Total Aplicada", f"{area_total:,.1f} Ha")

    # Gráfico de tendencia
    if 'FECHA' in super_base.columns and 'COSTO_NUM' in super_base.columns:
        tendencia = super_base.groupby(super_base['FECHA'].dt.month)['COSTO_NUM'].mean().reset_index()
        tendencia['MES'] = tendencia['FECHA'].map({1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'})
        fig = px.line(tendencia, x='MES', y='COSTO_NUM', markers=True, title="Tendencia de Costos por Mes")
        st.plotly_chart(fig, use_container_width=True)

if __name__ == "__main__":
    main()
