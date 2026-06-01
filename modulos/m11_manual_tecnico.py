import streamlit as st
import pandas as pd
import gspread
import io
import re
from datetime import datetime, timedelta

def a_numero_limpio(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip().replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1:
            partes = v.rsplit('.', 1)
            v = partes[0].replace('.', '') + '.' + partes[1]
        return float(v) if v else 0.0
    except: return 0.0

def generar_reporte_semanal_limpio(url_maestra, pestaña_nombre="TABLA 1"):
    try:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"])) if "gcp_credentials" in st.secrets else gspread.service_account(filename='credenciales.json')
        sh = gc.open_by_url(url_maestra)
        worksheet = sh.worksheet(pestaña_nombre)
        
        data = worksheet.get_all_values()
        if not data or len(data) < 6: return pd.DataFrame()
        
        encabezados = [str(c).upper().strip() for c in data[4]]
        filas_datos = data[5:]
        
        df = pd.DataFrame(filas_datos, columns=encabezados)
        
        df['FECHA_DT'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['FECHA_DT'])
        
        # Filtro de los últimos 7 días operativos
        fecha_limite = datetime.now() - timedelta(days=7)
        df_semanal = df[df['FECHA_DT'] >= fecha_limite].copy()
        
        if df_semanal.empty: return pd.DataFrame()
        
        df_semanal['FECHA'] = df_semanal['FECHA_DT'].dt.strftime('%d/%m/%Y')
        
        # Selección e integridad de columnas para el reporte externo de la empresa
        columnas_validas = [col for col in df_semanal.columns if any(c in col for c in ['ORDEN', 'BLOQUE', 'FINCA', 'SECTOR', 'BRUTA', 'FUMIG', 'COCTEL', 'FECHA', 'SEM', 'PILOTO', 'MODELO', 'PISTA'])]
        df_reporte = df_semanal[columnas_validas].copy()
        df_reporte.columns = [c.replace('\n', ' ').strip() for c in df_reporte.columns]
        
        return df_reporte
    except Exception as e:
        st.error(f"🚨 Error en la extracción satelital: {str(e)}")
        return pd.DataFrame()

def ejecutar(*args, **kwargs):
    st.markdown("<h1 style='text-align: center; color: #002244;'>📜 Módulo 11: Manual de Gobierno Técnico</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-style: italic; color: #64748b;'>Bóveda Institucional de Criterios y Seguridad Operativa</p>", unsafe_allow_html=True)
    
    # --- 📤 SECCIÓN DE DESPACHO INMEDIATA (PRIMERA LÍNEA DE FUEGO) ---
    st.markdown("---")
    st.markdown("### 📤 Extractor de Datos Seguro para la Empresa (Reporte Semanal)")
    st.write(
        "Este motor extrae la información consolidada de las misiones de los **últimos 7 días** "
        "directamente desde la TABLA 1. Genera un Excel plano **libre de costos confidenciales y fórmulas "
        "maestras**, aislando el archivo central de manipulaciones externas."
    )
    
    url_archivo_maestro = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    
    if st.button("🚀 EJECUTAR ESCÁNER Y COMPILAR REPORTE EXCEL", type="primary", use_container_width=True):
        with st.spinner("Conectando con la TABLA 1 y purgando registros de la última semana..."):
            df_limpio = generar_reporte_semanal_limpio(url_archivo_maestro)
            
            if df_limpio is not None and not df_limpio.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_limpio.to_excel(writer, index=False, sheet_name='Reporte Semanal Operaciones')
                buffer.seek(0)
                
                nombre_archivo = f"Reporte_Semanal_Operaciones_{datetime.now().strftime('%Y%m%d')}.xlsx"
                st.success(f"✅ **¡EXTRACCIÓN EXITOSA!** Se compilaron {len(df_limpio)} registros operativos listos para el despacho.")
                
                st.download_button(
                    label="📥 DESCARGAR EXCEL PLANO PARA ENVIAR A LA EMPRESA",
                    data=buffer,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ **ATENCIÓN:** El motor no detectó operaciones o misiones registradas en los últimos 7 días dentro del Archivo Maestro.")
                
    # --- MARCO TEÓRICO COMPLEMENTARIO ABAJO ---
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("### 🔬 Sustento de Gobierno del Sistema (Regla de Oro)")
    with st.expander("Ver especificaciones de Arquitectura BI"):
        st.write(
            "Para evitar el *Efecto Bolsa de Fechas*, los cálculos de frecuencias (Ciclos e Intervalos) "
            "se procesan finca por finca de manera aislada antes de emitir promedios. Esto blinda las métricas "
            "contra distorsiones artificiales cuando se selecciona la opción 'TODAS'."
        )
        st.latex(r"\text{Intervalo Promedio Zona} = \frac{\sum_{i=1}^{n} \text{Intervalo Finca}_i}{n}")
        
        # Generador de manual txt complementario
        texto_manual_md = f"MEMORIA TÉCNICA OFICIAL\nEmitido: {datetime.now().strftime('%Y-%m-%d')}\nRegla de Oro: Activa"
        st.download_button("📥 DESCARGAR DOCUMENTACIÓN EN TXT", texto_manual_md, "Manual_Genesis_BI.txt")
