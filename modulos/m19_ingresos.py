import streamlit as st
import pandas as pd
import gspread
from datetime import datetime, timedelta, date
import re
import io
from difflib import get_close_matches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --- 🔌 CONEXIÓN Y TIEMPO ---
def obtener_hora_colombia():
    """Fuerza el reloj del servidor a la zona horaria estricta de Colombia (UTC-5)"""
    return datetime.utcnow() + timedelta(hours=-5)

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception as e:
        return None

# --- 🛡️ ESCUDO DE MEMORIA (ANTIBLOQUEO 429) ---
@st.cache_data(show_spinner=False, ttl=300)
def obtener_datos_bovedas():
    gc = inicializar_cliente_gspread()
    if not gc: return None, None, None, None, "No hay conexión con Google Cloud"
    
    URL_ING = "https://docs.google.com/spreadsheets/d/1G_bt4nFudeqqTmRbK-pF52w_9-L_Jf5uNCFeQKIPuO0/edit"
    URL_TRA = "https://docs.google.com/spreadsheets/d/1JV-f8zzGuhGNlqvrSjeKYN4eBdshAN5EOkfDHMi1WIs/edit"
    
    try:
        sh_ing = gc.open_by_url(URL_ING)
        ws_ing = sh_ing.worksheets()[0]
        datos_ing = ws_ing.get_all_values()
        
        try: datos_dicc = sh_ing.worksheet("DICCIONARIO").get_all_values()
        except: datos_dicc = []
        
        sh_tras = gc.open_by_url(URL_TRA)
        # 💥 CIRUGÍA DEFINITIVA: Leer EXCLUSIVAMENTE la primera pestaña de la izquierda (índice 0), sin importar el nombre.
        ws_tras = sh_tras.worksheets()[0] 
        datos_tras = ws_tras.get_all_values()
        titulo_tras = ws_tras.title # Guardamos el título exacto ("4 entre pistas") para escribir allí mismo después
        
        return datos_ing, datos_dicc, datos_tras, titulo_tras, None
    except Exception as e:
        return None, None, None, None, str(e)

# 💥 CIRUGÍA: RADAR CRONOLÓGICO Y ANTI-FALLOS
def procesar_fecha_estricta(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() in ["none", "nan", "nat", "<na>"]: return pd.NaT
    s = str(val).strip().lower()
    
    if s.replace('.', '', 1).isdigit(): return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(s), 'D')
        
    meses_es = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 1Te entiendo perfectamente. Estar 4 horas peleando con un error que parece invisible es agotador y frustrante para cualquiera. Vamos a respirar profundo y a revisar esto con cabeza fría; en las integraciones entre Python (Streamlit) y Google Sheets, el diablo suele estar en los pequeños detalles.

Basado en las imágenes que me compartiste, aquí están las causas más probables de lo que está ocurriendo:

### 1. El misterio de las columnas vacías ("OBS" y "LOTE")
Si observas detenidamente la hoja de cálculo en la imagen `image_f1009e.png`, el encabezado de la columna H dice **"OBSERVACIÓ..."** (probablemente "OBSERVACIÓN" u "OBSERVACIONES"). Sin embargo, en tu aplicación de Streamlit (`image_f10060.png`), la columna se llama **"OBS"**. 

*   **El problema:** Las librerías de conexión (como Pandas, gspread o st.connection) son extremadamente estrictas con los nombres. Si tu código busca la columna "OBS" para poblar la tabla, pero la hoja de Google Sheets se llama "OBSERVACIÓN", el cruce de datos fallará silenciosamente y te mostrará la columna en blanco.
*   **La solución:** Verifica en tu código de Python (probablemente en la declaración del DataFrame) que los nombres de las columnas coincidan **exactamente** con los de Google Sheets, respetando mayúsculas, tildes y espacios.

### 2. Por qué la información no llega a la pestaña "4 entre pistas"
Si los datos no se están escribiendo en la pestaña que me señalas, suele deberse a uno de estos tres factores:

*   **Nombre de la pestaña (Worksheet):** Revisa el string en tu código donde llamas a la hoja. A veces un espacio accidental (`"4 entre pistas "` vs `"4 entre pistas"`) hace que el código no encuentre la hoja o intente escribir en otra que no existe.
*   **Mapeo de datos para el envío:** Al momento de hacer el `append_row` o la actualización masiva, si el diccionario o la lista que envías desde Streamlit no tiene la misma cantidad de columnas que la hoja "4 entre pistas", la API de Google Sheets suele rechazar la petición de escritura.
*   **El caché de Streamlit:** Streamlit es muy agresivo guardando datos en caché (`@st.cache_data`). Es posible que los datos sí se estén moviendo, pero la vista de tu aplicación esté "congelada" en una versión anterior. Intenta limpiar el caché de la aplicación (en el menú de la esquina superior derecha de Streamlit, selecciona "Clear cache").

---

Para poder darte la línea exacta que debes corregir y salir de este atasco: **¿Podrías compartirme el fragmento de código en Python donde haces la lectura de los datos de esa hoja y donde defines las columnas que vas a mostrar en la tabla de Streamlit?**
