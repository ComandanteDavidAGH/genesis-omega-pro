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

# --- 🖨️ MOTOR EXTRACTOR ADAPTABLE (HISTÓRICO O SEMANAL) ---
def generar_reporte_filtrado(url_maestra, filtrar_semana=False, pestaña_nombre="TABLA 1"):
    try:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"])) if "gcp_credentials" in st.secrets else gspread.service_account(filename='credenciales.json')
        sh = gc.open_by_url(url_maestra)
        worksheet = sh.worksheet(pestaña_nombre)
        
        data = worksheet.get_all_values()
        if not data or len(data) < 6: return pd.DataFrame()
        
        # Lectura alineada a la Fila 5 de su matriz
        encabezados = [str(c).upper().strip() for c in data[4]]
        filas_datos = data[5:]
        
        df = pd.DataFrame(filas_datos, columns=encabezados)
        
        # 🎯 FILTRO FRANCOTIRADOR: Columnas autorizadas para la empresa (Protección de costos y fórmulas)
        columnas_validas = [col for col in df.columns if any(c in col for c in ['ORDEN', 'BLOQUE', 'FINCA', 'SECTOR', 'BRUTA', 'FUMIG', 'COCTEL', 'FECHA', 'SEM', 'PILOTO', 'MODELO', 'PISTA'])]
        df_filtrado = df[columnas_validas].copy()
        df_filtrado.columns = [c.replace('\n', ' ').strip() for c in df_filtrado.columns]
        
        # Si la opción es solo la información semanal, aplicamos el filtro de fecha
        if filtrar_semana:
            df_filtrado['FECHA_DT'] = pd.to_datetime(df_filtrado['FECHA'], dayfirst=True, errors='coerce')
            df_filtrado = df_filtrado.dropna(subset=['FECHA_DT'])
            
            fecha_limite = datetime.now() - timedelta(days=7)
            df_filtrado = df_filtrado[df_filtrado['FECHA_DT'] >= fecha_limite].copy()
            
            if df_filtrado.empty: return pd.DataFrame()
            
            df_filtrado['FECHA'] = df_filtrado['FECHA_DT'].dt.strftime('%d/%m/%Y')
            df_filtrado = df_filtrado.drop(columns=['FECHA_DT'], errors='ignore')
            
        return df_filtrado
    except Exception as e:
        st.error(f"🚨 Error en el procesamiento de datos: {str(e)}")
        return pd.DataFrame()

# --- 📡 INTERFAZ LINEAL CORPORATIVA ---
def ejecutar(*args, **kwargs):
    st.markdown("<h1 style='text-align: center; color: #002244;'>📜 Módulo 11: Manual de Gobierno Técnico</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-style: italic; color: #64748b;'>Bóveda de Criterios Científicos y Seguridad de Despacho</p>", unsafe_allow_html=True)
    
    # 📥 SECCIÓN DE DESPACHO LOCAL (SOBERANÍA TOTAL)
    st.markdown("---")
    st.markdown("### 📤 Extractor de Datos Seguro para la Empresa")
    st.write(
        "Utilice estos dos controles de mando tácticos para descargar la información estructurada. "
        "Ambos archivos están **100% limpios de costos financieros confidenciales y fórmulas nativas de origen**, "
        "permitiéndole entregar reportes planos e impecables a la gerencia externa con su sello de calidad corporativo."
    )
    
    url_archivo_maestro = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📂 Operación Inicial (Todo el Histórico)")
        st.caption("Ideal para enviar por primera vez o restablecer auditorías completas.")
        if st.button("🚀 COMPILAR HISTÓRICO COMPLETO", key="btn_historico", use_container_width=True):
            with st.spinner("Descargando matriz completa y purificando columnas..."):
                df_hist = generar_reporte_filtrado(url_archivo_maestro, filtrar_semana=False)
                if not df_hist.empty:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_hist.to_excel(writer, index=False, sheet_name='Histórico Operaciones')
                    buffer.seek(0)
                    
                    st.success(f"✅ Compilados {len(df_hist)} registros históricos.")
                    st.download_button(
                        label="📥 DESCARGAR EXCEL MAESTRO PLANO",
                        data=buffer,
                        file_name=f"Reporte_Historico_Operaciones_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ No se encontraron datos en la TABLA 1.")
                    
    with col2:
        st.markdown("#### 📅 Operación Rutinaria (Últimos 7 Días)")
        st.caption("Ideal para alimentaciones semanales fijas de la empresa.")
        if st.button("⚡ COMPILAR INFORMACIÓN SEMANAL", key="btn_semanal", type="primary", use_container_width=True):
            with st.spinner("Aislando misiones de los últimos 7 días operativos..."):
                df_sem = generar_reporte_filtrado(url_archivo_maestro, filtrar_semana=True)
                if not df_sem.empty:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_sem.to_excel(writer, index=False, sheet_name='Reporte Semanal')
                    buffer.seek(0)
                    
                    st.success(f"✅ Purgadas {len(df_sem)} misiones de esta semana.")
                    st.download_button(
                        label="📥 DESCARGAR EXCEL SEMANAL",
                        data=buffer,
                        file_name=f"Reporte_Semanal_Operaciones_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ No se detectaron misiones en los últimos 7 días dentro de la TABLA 1.")

    # 🔬 MARCO DOCUMENTAL Y TEÓRICO EN LA PARTE INFERIOR (COMPLETO DE 4 APARTADOS)
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
        st.success("🎯 **Umbral Estructural de Ruptura:** Configurado en **> 5 días** de inactividad por lote.")

    with st.expander("📋 2. Diccionario de Variables Estables (Mapeo de Francotirador)"):
        st.write("Mapeo de dependencias analíticas configuradas para la estabilidad de la auditoría de misiones:")
        datos_diccionario = [
            {"Variable del Sistema": "FINCA_MAESTRA", "Origen en Matriz (Excel)": "Columna C (FINCA)", "Propósito Operacional": "Segmentación estricta de ciclos agrícolas por lote."},
            {"Variable del Sistema": "COSTO_MAESTRO", "Origen en Matriz (Excel)": "Columna W (VALOR FACTURAR)", "Propósito Operativo": "Cálculo real de la Media analítica de eficiencia financiera."},
            {"Variable del Sistema": "AREA_MAESTRA", "Origen en Matriz (Excel)": "Columna F (ÁREA FUMIG.)", "Propósito Operativo": "Sumatoria neta de hectáreas aplicadas sin duplicidad de compuestos."}
        ]
        st.table(pd.DataFrame(datos_diccionario))

    with st.expander("⚙️ 3. Lógica del Algoritmo Temporal y Segmentación"):
        st.write(
            "El sistema procesa los deltas cronológicos utilizando objetos indexados en memoria. "
            "Al presionar el escáner, los datos sufren una transformación matemática limpia:"
        )
        st.markdown(
            "* **Paso A:** Conversión de cadenas de texto de Google Sheets a formatos numéricos puros de coma flotante.\n"
            "* **Paso B:** Purificación de duplicados operativos basados en la terna invariable (Fecha, Finca, Número de OS).\n"
            "* **Paso C:** Medición de intervalos reales mediante cálculo vectorial directo de fechas."
        )

    with st.expander("📥 4. Biblioteca de Descarga de Manuales Oficiales"):
        st.write("Descargue las versiones en texto plano de respaldo técnico para auditorías externas o archivos físicos de la mesa de mando:")
        texto_manual_md = (
            "INFORME DE ARQUITECTURA TÉCNICA INSTITUCIONAL - GÉNESIS OMEGA PRO\n"
            f"Compilado el: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            "Regla de Oro: ACTIVADA Y BLINDADA\n"
            "Constante de ciclo: > 5 días\n"
            "Mapeo de extracción de datos: Columna W, Columna F y Columna G del archivo maestro."
        )
        st.download_button(
            label="📥 DESCARGAR MANUAL DE ARQUITECTURA EN TXT",
            data=texto_manual_md,
            file_name="Memoria_Tecnica_Completa_Genesis.txt",
            mime="text/plain",
            use_container_width=True
        )
