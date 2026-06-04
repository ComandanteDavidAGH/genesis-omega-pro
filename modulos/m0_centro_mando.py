import streamlit as st
import pandas as pd
import numpy as np

def renderizar():
    # 🚀 MOTOR VISUAL VIP: Estilizado del Centro de Mando e Inyección CSS
    st.markdown("""
    <style>
    .titulo-principal { 
        color: #0d1b2a; 
        border-bottom: 3px solid #d4af37; 
        padding-bottom: 5px; 
        font-family: 'Arial Black', sans-serif; 
    }
    
    /* Contenedores Oficiales para Tablas */
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
        border: 3px solid #0d1b2a !important;
        border-radius: 8px !important;
        box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important;
        overflow: hidden !important;
    }
    
    /* Mini-KPIs del Centro de Mando */
    .hud-mando {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-left: 5px solid #0d1b2a;
        padding: 12px 20px;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #dee2e6;
    }
    .hud-mando-item { text-align: center; }
    .hud-mando-title { font-size: 11px; color: #6c757d; font-family: 'Arial Black', sans-serif; text-transform: uppercase; margin: 0; }
    .hud-mando-value { font-size: 20px; color: #0d1b2a; font-weight: 900; margin: 0; }
    .hud-mando-alert { color: #cc0000; font-family: 'Arial Black', sans-serif; }
    .hud-mando-ok { color: #00994c; font-family: 'Arial Black', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>🏠 Centro de Mando y Control</h1>", unsafe_allow_html=True)
    
    # --- SALUDO OFICIAL ---
    st.info("📡 **Radar Principal:** Monitoreo activo de sistemas, escuadrones y logística aérea.")
    st.markdown(f"### Bienvenido al Cuartel General, **{st.session_state.get('usuario_nombre', 'Comandante')}**.")
    st.write("El sistema Génesis Omega Pro se encuentra en línea y operando bajo parámetros óptimos. Seleccione un hangar en el menú lateral para iniciar operaciones.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # --- 🚨 RADAR LOGÍSTICO DE ALERTA TEMPRANA ---
    st.markdown("### 🚨 Radar Logístico: Alerta Temprana de Inventarios")
    
    df_sabana = st.session_state.get('df_sabana', pd.DataFrame())
    
    if df_sabana.empty:
        st.warning("⚠️ **Radar en Modo Espera:** El sistema no detecta un inventario activo en la memoria. Para encender el radar, por favor cargue la **Sábana SAP** actualizada en el **📥 Módulo 2 (Carga Facturación)**.")
    else:
        with st.spinner("Sincronizando existencias de pistas y consolidando lotes..."):
            
            # 1. Mapeo estructural indexado según la Sábana de SAP real del usuario
            col_cod = next((c for c in df_sabana.columns if str(c).strip() == 'Material'), None)
            col_pista = next((c for c in df_sabana.columns if str(c).strip() == 'Almacén'), None)
            col_saldo = next((c for c in df_sabana.columns if str(c).strip() == 'Libre utilización'), None)
            col_desc = next((c for c in df_sabana.columns if str(c).strip() == 'Descripción del material'), None)

            # Fallbacks automáticos por si el formato cambia de mayúsculas/minúsculas
            if not col_cod: col_cod = next((c for c in df_sabana.columns if 'MATERIAL' in str(c).upper()), None)
            if not col_pista: col_pista = next((c for c in df_sabana.columns if 'ALMACEN' in str(c).upper() or 'LGORT' in str(c).upper()), None)
            if not col_saldo: col_saldo = next((c for c in df_sabana.columns if 'LIBRE' in str(c).upper() or 'UTILIZACION' in str(c).upper() or 'LABST' in str(c).upper()), None)
            if not col_desc: col_desc = next((c for c in df_sabana.columns if 'DESC' in str(c).upper() or 'TEXTO' in str(c).upper()), None)

            if not col_cod or not col_pista or not col_saldo:
                st.error("❌ Error de Radar: No se pudieron mapear las columnas. Verifique que el archivo corresponda a la Sábana Estándar.")
            else:
                # 2. Copia y Limpieza Vectorial de Registros Activos
                df_temp = df_sabana.copy()
                df_temp[col_saldo] = pd.to_numeric(df_temp[col_saldo].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                df_temp = df_temp[df_temp[col_saldo] > 0]
                
                # Formatear el código quitando decimales flotantes (.0) de Excel
                codigos_limpios = df_temp[col_cod].astype(str).str.split('.').str[0].str.strip()
                
                # 🧪 FUSIÓN LOGÍSTICA ATÓMICA: "CÓDIGO | NOMBRE REAL DEL PRODUCTO"
                if col_desc:
                    df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | " + df_temp[col_desc].astype(str).str.strip().str.upper()
                else:
                    df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | INSUMO QUÍMICO REGISTRADO"

                # 3. CONSOLIDACIÓN MAESTRA: Suma todos los lotes de una misma pista y producto
                inventario_agrupado = df_temp.groupby([col_pista, 'PRODUCTO_RADAR'])[col_saldo].sum().reset_index()
                
                # 4. COMPILADOR DE REGLAS DE SEGURIDAD OPERATIVA
                pistas_series = inventario_agrupado[col_pista].astype(str).str.upper()
                productos_series = inventario_agrupado['PRODUCTO_RADAR'].astype(str).str.upper()
                
                es_pista_menor = pistas_series.str.contains("LUCI|TEHO", na=False)
                es_aceite = productos_series.str.contains("ACEITE|GRANEL|COMBUSTIBLE|DICAM", na=False)
                es_mancol = productos_series.str.contains("MANCOL|MANCOZEB|103680|104287", na=False)
                es_aditivo = productos_series.str.contains("ACONDICIONADOR|NATURAMIN|105980|108214|105296", na=False)
                
                condiciones = [
                    es_aceite & es_pista_menor,
                    es_aceite & ~es_pista_menor,
                    es_mancol & es_pista_menor,
                    es_mancol & ~es_pista_menor,
                    es_aditivo
                ]
                
                valores_limite = [1000, 30280, 1000, 2500, 30]
                regles_texto = [
                    "1.000 L (Aceite - Pista Menor)",
                    "30,280 L (Aceite - Pista Principal)",
                    "1,000 L (Mancol - Pista Menor)",
                    "2,500 L (Mancol - Pista Principal)",
                    "30 L/Kg (Aditivo de Alta Rotación)"
                ]
                
                inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD'] = np.select(condiciones, valores_limite, default=100)
                inventario_agrupado['📋 REGLA APLICADA'] = np.select(condiciones, regles_texto, default="100 L/Kg (Estándar Global)")
                
                # 5. Filtrar solo los elementos que rompieron el stock crítico
                df_alertas = inventario_agrupado[inventario_agrupado[col_saldo] < inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD']].copy()
                
                df_alertas = df_alertas.rename(columns={
                    col_pista: "📍 PISTA / ALMACÉN",
                    'PRODUCTO_RADAR': "🧪 CÓDIGO | NOMBRE DEL PRODUCTO",
                    col_saldo: "⚠️ SALDO ACTUAL"
                })
                
                columnas_finales = ["📍 PISTA / ALMACÉN", "🧪 CÓDIGO | NOMBRE DEL PRODUCTO", "⚠️ SALDO ACTUAL", "🛡️ LÍMITE DE SEGURIDAD", "📋 REGLA APLICADA"]
                df_alertas_render = df_alertas[columnas_finales].sort_values(by="📍 PISTA / ALMACÉN")
                
                # 6. HUD TÁCTICO CORREGIDO (Ahora sí leerá las 5 pistas por separado)
                total_almacenes = inventario_agrupado[col_pista].nunique()
                total_insumos = inventario_agrupado['PRODUCTO_RADAR'].nunique()
                conteo_alertas = len(df_alertas_render)
                
                clase_alerta = "hud-mando-value hud-mando-alert" if conteo_alertas > 0 else "hud-mando-value hud-mando-ok"
                texto_alerta = f"{conteo_alertas} Alertas" if conteo_alertas > 0 else "0 Críticos"
                
                st.markdown(f"""
                <div class="hud-mando">
                    <div class="hud-mando-item">
                        <p class="hud-mando-title">Pistas / Almacenes Activos</p>
                        <p class="hud-mando-value">🛰️ {total_almacenes}</p>
                    </div>
                    <div class="hud-mando-item">
                        <p class="hud-mando-title">Insumos Consolidados Únicos</p>
                        <p class="hud-mando-value">🧪 {total_insumos}</p>
                    </div>
                    <div class="hud-mando-item">
                        <p class="hud-mando-title">Estado de Carga</p>
                        <p class="{clase_alerta}">{texto_alerta}</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # 7. Renderizado Final de la Tabla Estilizada
                if conteo_alertas > 0:
                    st.error(f"🚨 **¡ALERTA ROJA! MÁRGENES OPERATIVOS CRÍTICOS DETECTADOS:**")
                    
                    df_alertas_render["⚠️ SALDO ACTUAL"] = df_alertas_render["⚠️ SALDO ACTUAL"].apply(lambda x: f"{x:,.1f}".replace(",", "."))
                    df_alertas_render["🛡️ LÍMITE DE SEGURIDAD"] = df_alertas_render["🛡️ LÍMITE DE SEGURIDAD"].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                    
                    def pintar_rojo_elegante(val):
                        return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold; border-bottom: 1px solid #dee2e6;'] * len(val)
                        
                    st.dataframe(df_alertas_render.style.apply(pintar_rojo_elegante, axis=1), use_container_width=True, hide_index=True)
                else:
                    st.success("✅ **INVENTARIO ÓPTIMO:** Todos los insumos químicos y energéticos en la totalidad de las pistas se encuentran por encima de los márgenes de seguridad establecidos. Operación aérea asegurada.")
