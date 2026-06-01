import streamlit as st
import pandas as pd
import gspread
import io
import re
from datetime import datetime, timedelta

# --- 🧪 TRADUCTOR SEGURO DE NÚMEROS ---
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

# --- 🖨️ MOTOR EXCRIPTOR DE REPORTE PLANO ---
def generar_reporte_semanal_limpio(url_maestra, pestaña_nombre="TABLA 1"):
    try:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"])) if "gcp_credentials" in st.secrets else gspread.service_account(filename='credenciales.json')
        sh = gc.open_by_url(url_maestra)
        worksheet = sh.worksheet(pestaña_nombre)
        
        data = worksheet.get_all_values()
        if not data or len(data) < 6: return pd.DataFrame()
        
        # Lectura estricta alineada a la Fila 5 del Comandante
        encabezados = [str(c).upper().strip() for c in data[4]]
        filas_datos = data[5:]
        
        df = pd.DataFrame(filas_datos, columns=encabezados)
        df['FECHA_DT'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['FECHA_DT'])
        
        # Filtro temporal estricto (Últimos 7 días de operación)
        fecha_limite = datetime.now() - timedelta(days=7)
        df_semanal = df[df['FECHA_DT'] >= fecha_limite].copy()
        
        if df_semanal.empty: return pd.DataFrame()
        
        df_semanal['FECHA'] = df_semanal['FECHA_DT'].dt.strftime('%d/%m/%Y')
        
        # 🎯 FILTRO FRANCOTIRADOR: Columnas autorizadas para la empresa (Sin dinero interno)
        columnas_validas = [col for col in df_semanal.columns if any(c in col for c in ['ORDEN', 'BLOQUE', 'FINCA', 'SECTOR', 'BRUTA', 'FUMIG', 'COCTEL', 'FECHA', 'SEM', 'PILOTO', 'MODELO', 'PISTA'])]
        df_reporte = df_semanal[columnas_validas].copy()
        df_reporte.columns = [c.replace('\n', ' ').strip() for c in df_reporte.columns]
        
        return df_reporte
    except Exception as e:
        st.error(f"🚨 Error en el procesamiento del reporte: {str(e)}")
        return pd.DataFrame()

# --- 📡 INTERFAZ LINEAL CORPORATIVA ---
def ejecutar(*args, **kwargs):
    st.markdown("<h1 style='text-align: center; color: #002244;'>📜 Módulo 11: Manual de Gobierno Técnico</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-style: italic; color: #64748b;'>Bóveda de Criterios Científicos y Seguridad de Despacho</p>", unsafe_allow_html=True)
    
    # 📥 SECCIÓN DE DESPACHO LOCAL (SOBERANÍA TOTAL)
    st.markdown("---")
    st.markdown("### 📤 Extractor de Datos Seguro para la Empresa")
    st.write(
        "Presione el botón para compilar la información de los **últimos 7 días**. "
        "El sistema generará un archivo de Excel plano descargable en su computador, **limpio de costos "
        "confidenciales y fórmulas de origen**, listo para que usted lo guarde o envíe por sus canales oficiales."
    )
    
    url_archivo_maestro = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    
    if st.button("🚀 EJECUTAR ESCÁNER Y COMPILAR REPORTE EXCEL", type="primary", use_container_width=True):
        with st.spinner("Escaneando la TABLA 1 y aislando registros de la última semana..."):
            df_limpio = generar_reporte_semanal_limpio(url_archivo_maestro)
            
            if df_limpio is not None and not df_limpio.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_limpio.to_excel(writer, index=False, sheet_name='Reporte Semanal Operaciones')
                buffer.seek(0)
                
                # Nombre de archivo dinámico con la fecha de hoy
                nombre_archivo = f"Reporte_Semanal_Operaciones_{datetime.now().strftime('%Y%m%d')}.xlsx"
                st.success(f"✅ **¡EXTRACCIÓN FILTRADA CON ÉXITO!** Se aislaron {len(df_limpio)} misiones operativas.")
                
                st.download_button(
                    label="📥 DESCARGAR EXCEL PLANO SEGURO",
                    data=buffer,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ **ATENCIÓN:** El motor no detectó misiones u operaciones registradas en los últimos 7 días dentro de la TABLA 1.")

    # 🔬 MARCO DOCUMENTAL Y TEÓRICO EN LA PARTE INFERIOR
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("### 🔬 Núcleo Teórico y Sustento del Sistema")
    
    with st.expander("🔬 1. Principios Operativos y la Regla de Oro"):
        st.markdown("#### Principio de Aislamiento por Propiedad")
        st.write(
            "Para mitigar la distorsión analítica del *Efecto Bolsa de Fechas*, los cálculos de frecuencias "
            "(Ciclos e Intervalos) se procesan finca por finca de manera aislada antes de emitir promedios globales. "
            "Esto blinda las métricas corporativas contra variaciones artificiales cuando se evalúa la opción 'TODAS'."
        )
        st.latex(r"\text{Intervalo Promedio Zona} = \frac{\sum_{i=1}^{n} \text{Intervalo Finca}_i}{n}")
        st.success("🎯 **Umbral Estructural de Ruptura:** Configurado en **> 5 días** de inactividad.")

    with st.expander("📋 2. Diccionario de Variables Estables"):
        st.write("Mapeo de dependencias analíticas configuradas para la auditoría de misiones:")
        datos_diccionario = [
            {"Variable del Sistema": "FINCA_MAESTRA", "Origen en Matriz (Excel)": "Columna C (FINCA)", "Propósito Operacional": "Segmentación estricta de ciclos agrícolas por lote."},
            {"Variable del Sistema": "COSTO_MAESTRO", "Origen en Matriz (Excel)": "Columna W (VALOR FACTURAR)", "Propósito Operativo": "Cálculo real de la Media analítica de eficiencia financiera."},
            {"Variable del Sistema": "AREA_MAESTRA", "Origen en Matriz (Excel)": "Columna F (ÁREA FUMIG.)", "Propósito Operativo": "Sumatoria neta de hectáreas aplicadas sin duplicidad de compuestos."}
        ]
        st.table(pd.DataFrame(datos_diccionario))

    with st.expander("📥 3. Biblioteca de Documentación Corporativa"):
        st.write("Descargue los respaldos oficiales de la arquitectura del software:")
        texto_manual_md = f"MEMORIA TÉCNICA MAESTRA - AGROAÉREO TÁCTICO\nEmitido de forma segura: {datetime.now().strftime('%Y-%m-%d %H:%M')}\nEstatus de la Regla de Oro: PROTEGIDA Y ACTIVA"
        st.download_button("📥 DESCARGAR MANUAL TÉCNICO (.TXT)", texto_manual_md, "Manual_Genesis_BI.txt", use_container_width=True)
