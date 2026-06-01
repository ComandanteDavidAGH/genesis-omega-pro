import streamlit as st
import pandas as pd
import gspread
import io
import re
from datetime import datetime, timedelta

# --- 🧪 TRADUCTOR NEUTRO PARA EL REPORTE SEMANAL ---
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
        # Conexión satelital usando las credenciales seguras del sistema
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"])) if "gcp_credentials" in st.secrets else gspread.service_account(filename='credenciales.json')
        sh = gc.open_by_url(url_maestra)
        worksheet = sh.worksheet(pestaña_nombre)
        
        data = worksheet.get_all_values()
        if not data or len(data) < 6: return pd.DataFrame()
        
        # Mapeo exacto según la radiografía de la Fila 5 (índice 4)
        encabezados = [str(c).upper().strip() for c in data[4]]
        filas_datos = data[5:]
        
        df = pd.DataFrame(filas_datos, columns=encabezados)
        
        # Filtrado temporal inteligente (últimos 7 días)
        df['FECHA_DT'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['FECHA_DT'])
        
        fecha_limite = datetime.now() - timedelta(days=7)
        df_semanal = df[df['FECHA_DT'] >= fecha_limite].copy()
        
        if df_semanal.empty: return pd.DataFrame()
        
        # Restauramos el formato legible de fecha antes del despacho
        df_semanal['FECHA'] = df_semanal['FECHA_DT'].dt.strftime('%d/%m/%Y')
        
        # 🎯 FILTRO LÁSER: Solo las columnas requeridas por la empresa (Cero fórmulas, cero costos confidenciales)
        columnas_empresa = [
            'Nº ORDEN', 'BLOQUE', 'FINCA', 'SECTOR', 
            'ÁREA BRUTA\n(HA)', 'ÁREA FUMIG.\n(HA)', 'COCTEL', 
            'FECHA', 'DÌA SEM', 'SEM', 'PILOTO', 'MODELO', 'PISTA'
        ]
        
        # Validamos nombres normalizados para evitar caídas
        columnas_validas = [col for col in df_semanal.columns if any(c in col for c in ['ORDEN', 'BLOQUE', 'FINCA', 'SECTOR', 'BRUTA', 'FUMIG', 'COCTEL', 'FECHA', 'SEM', 'PILOTO', 'MODELO', 'PISTA'])]
        df_reporte = df_semanal[columnas_validas].copy()
        
        # Limpieza de textos crudos de saltos de línea para presentación ejecutiva
        df_reporte.columns = [c.replace('\n', ' ').strip() for c in df_reporte.columns]
        return df_reporte

    except Exception as e:
        st.error(f"🚨 Error en la extracción satelital: {str(e)}")
        return pd.DataFrame()

# --- 📡 INTERFAZ COMPLETA DEL MANUAL Y DESPACHO ---
def ejecutar(*args, **kwargs):
    st.markdown("<h1 style='text-align: center; color: #002244;'>📜 Módulo 11: Manual de Gobierno Técnico</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-style: italic; color: #64748b;'>Bóveda Institucional de Criterios y Seguridad Operativa</p>", unsafe_allow_html=True)
    
    st.info("🎓 **NÚCLEO TEÓRICO COMPILADO:** Este espacio custodia las leyes matemáticas de la plataforma y el puente seguro de exportación de datos.")

    # Pestañas del módulo
    tab_principios, tab_diccionario, tab_despacho = st.tabs([
        "🔬 1. Principios y Regla de Oro",
        "📋 2. Diccionario de Variables",
        "📤 3. Despacho de Reporte Semanal"
    ])

    # --- PESTAÑA 1: PRINCIPIOS Y REGLA DE ORO ---
    with tab_principios:
        st.markdown("### 🏛️ La Regla de Oro: Principio de Aislamiento Operativo por Propiedad")
        st.write(
            "Para mitigar la distorsión analítica del *Efecto Bolsa de Fechas*, el motor BI tiene "
            "prohibido unificar registros cronológicos antes de su procesamiento. El sistema calcula los ciclos "
            "e intervalos reales finca por finca de manera independiente, y consolida la zona mediante una media aritmética pura."
        )
        st.markdown("#### Fórmulas de Control:")
        st.latex(r"\text{Intervalo Promedio Zona} = \frac{\sum_{i=1}^{n} \text{Intervalo Finca}_i}{n}")
        st.success("✅ **Umbral de Ruptura Operacional:** Configurado rígidamente en **> 5 días** de inactividad por propiedad.")

    # --- PESTAÑA 2: DICCIONARIO DE VARIABLES ---
    with tab_diccionario:
        st.markdown("### 🎯 Especificación de Cabeceras Estables")
        datos_diccionario = [
            {"Variable Interna": "FINCA_MAESTRA", "Origen Raíz (Excel)": "Columna B (FINCA)", "Uso": "Aislamiento de bucles por propiedad."},
            {"Variable Interna": "COSTO_MAESTRO", "Origen Raíz (Excel)": "Columna W (VALOR A FACTURAR...)", "Uso": "Cálculo de la Media pura de costos por Ha."},
            {"Variable Interna": "AREA_MAESTRA", "Origen Raíz (Excel)": "Columna F (ÁREA FUMIG.)", "Uso": "Sumatoria de Ha operadas sin químicos repetidos."}
        ]
        st.table(pd.DataFrame(datos_diccionario))

    # --- PESTAÑA 3: DESPACHO DE REPORTE SEMANAL ---
    with tab_despacho:
        st.markdown("### 📤 Extractor de Datos Seguro para la Empresa")
        st.write(
            "Use este control operativo para extraer la información consolidada de los **últimos 7 días**. "
            "El motor generará un archivo plano en Excel estructurado con los valores puros de las misiones terrestres "
            "y aéreas, **eliminando fórmulas sensibles, costos confidenciales y protegiendo el Drive Maestro contra modificaciones.**"
        )
        
        url_archivo_maestro = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
        
        if st.button("🚀 EJECUTAR ESCÁNER Y COMPILAR REPORTE", type="primary", use_container_width=True):
            with st.spinner("Descargando matriz y aislando registros de la última semana..."):
                df_limpio = generar_reporte_semanal_limpio(url_archivo_maestro)
                
                if df_limpio is not None and not df_limpio.empty:
                    buffer = io.BytesIO()
                    # Generamos el Excel nativo plano listo para descarga
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_limpio.to_excel(writer, index=False, sheet_name='Reporte Semanal Operaciones')
                    buffer.seek(0)
                    
                    nombre_archivo = f"Reporte_Semanal_Operaciones_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    st.success(f"✅ **¡EXTRACCIÓN EXITOSA!** Se encontraron {len(df_limpio)} misiones operativas en los últimos 7 días.")
                    
                    st.download_button(
                        label="📥 DESCARGAR EXCEL PLANO PARA ENVIAR A LA EMPRESA",
                        data=buffer,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ **ATENCIÓN:** No se detectaron misiones u operaciones registradas en los últimos 7 días dentro de la TABLA 1.")
