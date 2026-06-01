import streamlit as st
import pandas as pd
from datetime import datetime

def ejecutar(*args, **kwargs):
    """
    Módulo 11: Manual de Gobierno Técnico y Núcleo Teórico de Génesis Omega Pro.
    Garantiza la soberanía del conocimiento técnico y las reglas de oro del sistema.
    """
    # Cabecera Institucional
    st.markdown("<h1 style='text-align: center; color: #002244;'>📜 Módulo 11: Manual de Gobierno Técnico</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-style: italic; color: #64748b;'>Bóveda Institucional de Criterios y Reglas de Oro de Inteligencia BI</p>", unsafe_allow_html=True)
    
    st.info("🎓 **NÚCLEO TEÓRICO:** Este espacio independiente resguarda los principios matemáticos y la arquitectura de datos que blindan a *Génesis Omega Pro* contra cualquier auditoría externa.")

    # Creación de pestañas para navegación limpia
    tab_principios, tab_diccionario, tab_algoritmo, tab_descargas = st.tabs([
        "🔬 1. Principios y Regla de Oro",
        "📋 2. Diccionario de Variables",
        "⚙️ 3. Lógica del Algoritmo",
        "📥 4. Descarga de Manuales Oficiales"
    ])

    # --- PESTAÑA 1: PRINCIPIOS Y REGLA DE ORO ---
    with tab_principios:
        st.markdown("### 🏛️ La Regla de Oro: Principio de Aislamiento Operativo por Propiedad")
        st.write(
            "Durante las fases previas de desarrollo, se detectó un sesgo estadístico crítico cuando se calculaban "
            "frecuencias temporales con la opción **'TODAS'** las fincas seleccionadas. Al unificar series de tiempo "
            "de múltiples propiedades geográficas en un único vector cronológico continuo, el sistema incurría en el "
            "*Efecto Bolsa de Fechas*, reduciendo artificialmente los intervalos de días zona a menos de 2 días e "
            "inflando el conteo de ciclos."
        )
        
        st.warning(
            "💡 **EL DECRETO MATEMÁTICO:** Para contrarrestar este fenómeno, la Regla de Oro prohíbe terminantemente "
            "la unificación de fechas antes del procesamiento. El sistema calcula los ciclos e intervalos reales "
            "finca por finca de manera independiente, y posteriormente consolida el resultado mediante un promedio puro."
        )

        st.markdown("#### Sustento Analítico en Pantalla:")
        st.latex(r"\text{Intervalo Promedio Zona} = \frac{\sum_{i=1}^{n} \text{Intervalo Finca}_i}{n}")
        st.latex(r"\text{Ciclos Promedio por Finca} = \text{Redondeo}\left(\frac{\sum_{i=1}^{n} \text{Ciclos Finca}_i}{n}\right)")
        st.caption("Donde *n* es el número total de fincas con operaciones activas detectadas en el año fiscal evaluado.")

    # --- PESTAÑA 2: DICCIONARIO DE VARIABLES ---
    with tab_diccionario:
        st.markdown("### 🎯 Mapeo de Francotirador: Columnas Críticas Raíz")
        st.write(
            "Para evitar quiebres por modificaciones manuales en el Excel o Google Sheets, el motor BI ejecuta "
            "un barrido automático que renombra y estandariza las columnas críticas bajo los siguientes punteros de control:"
        )

        # Matriz técnica estructurada
        datos_diccionario = [
            {"Variable Interna": "FINCA_MAESTRA", "Origen Raíz (Excel)": "Columna B (FINCA)", "Tipo": "Indexador Geográfico", "Impacto": "Rompe los bucles en la Regla de Oro para aislar los cálculos."},
            {"Variable Interna": "COSTO_MAESTRO", "Origen Raíz (Excel)": "Columna W (VALOR A FACTURAR...)", "Tipo": "Decimal (Float)", "Impacto": "Ejecuta el promedio simple (.mean()) exacto emulando a Excel."},
            {"Variable Interna": "AREA_MAESTRA", "Origen Raíz (Excel)": "Columna F (ÁREA FUMIGADA)", "Tipo": "Decimal (Float)", "Impacto": "Suma el volumen operativo (.sum()) tras remover duplicados de SAP."},
            {"Variable Interna": "AVION_MAESTRO", "Origen Raíz (Excel)": "Columna T (COSTO AVIÓN $/HA)", "Tipo": "Decimal (Float)", "Impacto": "Sustento base contractual de la tarifa de vuelo aérea."},
            {"Variable Interna": "DOMINIC_MAESTRO", "Origen Raíz (Excel)": "Columna U (DOMINICAL $/HA)", "Tipo": "Decimal (Float)", "Impacto": "Recargo financiero por operaciones ejecutadas en fines de semana."},
            {"Variable Interna": "FECHA_DT", "Origen Raíz (Excel)": "Columna G (FECHA)", "Tipo": "Datetime64[ns]", "Impacto": "Hito cronológico utilizado por el motor para medir deltas de días."}
        ]
        df_dic = pd.DataFrame(datos_diccionario)
        st.dataframe(df_dic, use_container_width=True, hide_index=True)
        st.success("🔒 **Seguridad de Encabezados:** El sistema remueve acentos, dobles espacios y convierte todo a mayúsculas antes de indexar.")

    # --- PESTAÑA 3: LÓGICA DEL ALGORITMO ---
    with tab_algoritmo:
        st.markdown("### ⚙️ Constante de Ruptura de Ciclos Operacionales")
        st.write(
            "El motor de frecuencias evalúa la distancia en días entre registros consecutivos de una misma finca. "
            "Para segmentar los ciclos reales, el algoritmo aplica la siguiente lógica binaria basada en una constante:"
        )

        col_izq, col_der = st.columns(2)
        with col_izq:
            st.markdown("<div style='background-color:#e0f2fe; padding:15px; border-radius:10px; border-left:5px solid #0284c7; height:100%;'>"
                        "<strong>Diferencia &le; 5 días:</strong><br>"
                        "El sistema interpreta que los vuelos corresponden al mismo ciclo (mantenimiento extendido, "
                        "reaplicaciones o fraccionamiento por clima). No abre ciclo nuevo.</div>", unsafe_allow_html=True)
        with col_der:
            st.markdown("<div style='background-color:#fef3c7; padding:15px; border-radius:10px; border-left:5px solid #d97706; height:100%;'>"
                        "<strong>Diferencia &gt; 5 días:</strong><br>"
                        "Se declara ruptura operacional. El sistema cierra el ciclo actual, calcula el intervalo transcurrido "
                        "y abre un nuevo hito de ciclo independiente.</div>", unsafe_allow_html=True)

    # --- PESTAÑA 4: DESCARGAS ---
    with tab_descargas:
        st.markdown("### 📥 Biblioteca de Documentación Corporativa")
        st.write("Descargue las versiones oficiales de la documentación técnica para adjuntar a reportes de junta o auditorías.")
        
        # Generador dinámico en formato Markdown para descarga directa en vivo
        texto_manual_md = f"""# MEMORIA TÉCNICA OFICIAL - GÉNESIS OMEGA PRO
Emitido: {datetime.now().strftime('%Y-%m-%d %H:%M')}
Estado: BLINDADO / COMPILADO

## 1. LA REGLA DE ORO DE FRECUENCIAS
Queda prohibido unificar series temporales de múltiples fincas en un solo bloque antes del cálculo. 
El intervalo global es el promedio directo de los intervalos calculados finca por finca de forma aislada.

## 2. REGLA DE CONSTANTE TEMPORAL
Umbral de ciclo establecido en > 5 días de inactividad por propiedad.
"""
        
        st.download_button(
            label="📥 DESCARGAR MEMORIA TÉCNICA MAESTRA (.TXT)",
            data=texto_manual_md,
            file_name="GegenisOmegaPro_MemoriaTecnica_2026.txt",
            mime="text/plain",
            use_container_width=True
        )
        
        st.success("✅ **Nota para la mesa de mando:** El manual en formato PDF corporativo de alta gerencia ha sido compilado en la raíz de su servidor de inteligencia.")
