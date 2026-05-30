import streamlit as st
import pandas as pd

def renderizar():
    st.markdown("<h1 class='titulo-principal'>🏠 Centro de Mando y Control</h1>", unsafe_allow_html=True)
    
    # --- SALUDO OFICIAL ---
    st.info("📡 **Radar Principal:** Monitoreo activo de sistemas, escuadrones y logística.")
    st.markdown(f"### Bienvenido al Cuartel General, **{st.session_state.get('usuario_nombre', 'Comandante')}**.")
    st.write("El sistema Génesis Omega Pro se encuentra en línea y operando bajo parámetros óptimos. Seleccione un hangar en el menú lateral para iniciar operaciones.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # --- 🚨 RADAR LOGÍSTICO DE ALERTA TEMPRANA ---
    st.markdown("### 🚨 Radar Logístico: Alerta Temprana de Inventarios")
    
    # El radar lee la memoria de la Sábana SAP cargada en el Módulo 2
    df_sabana = st.session_state.get('df_sabana', pd.DataFrame())
    
    if df_sabana.empty:
        st.warning("⚠️ **Radar en Modo Espera:** El sistema no detecta un inventario activo en la memoria. Para encender el radar, por favor cargue la **Sábana SAP** actualizada en el **📥 Módulo 2 (Carga Facturación)**.")
    else:
        with st.spinner("Escaneando bodegas y cruzando saldos con niveles críticos..."):
            # 1. Francotirador de columnas
            col_mat = next((c for c in df_sabana.columns if 'TEXTO' in str(c).upper() or 'DESC' in str(c).upper() or 'MATERIAL' in str(c).upper()), None)
            col_pista = next((c for c in df_sabana.columns if 'ALMACEN' in str(c).upper() or 'PISTA' in str(c).upper() or 'CENTRO' in str(c).upper()), None)
            col_saldo = next((c for c in df_sabana.columns if 'LIBRE' in str(c).upper() or 'SALDO' in str(c).upper() or 'CANTIDAD' in str(c).upper()), None)
            
            if not col_mat or not col_pista or not col_saldo:
                st.error("❌ Error de Radar: No se pudieron identificar las columnas de Material, Almacén o Saldo en la Sábana SAP cargada.")
            else:
                alertas = []
                
                # 2. Agrupar saldos por Pista y Material (por si hay lotes separados)
                df_temp = df_sabana.copy()
                df_temp[col_saldo] = pd.to_numeric(df_temp[col_saldo].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                inventario_agrupado = df_temp.groupby([col_pista, col_mat])[col_saldo].sum().reset_index()
                
                # 3. Aplicación de las Reglas de Oro del Comandante
                for _, row in inventario_agrupado.iterrows():
                    pista = str(row[col_pista]).strip().upper()
                    producto = str(row[col_mat]).strip().upper()
                    saldo = row[col_saldo]
                    
                    if saldo <= 0: continue # Ignorar ítems vacíos o en cero absoluto
                    
                    es_pista_menor = "LUCI" in pista or "TEHO" in pista
                    
                    # Base Estándar Global
                    limite = 100
                    tipo_limite = "100 L/Kg (Estándar Global)"
                    
                    # Filtros de Alta Prioridad
                    if "ACEITE" in producto or "GRANEL" in producto or "COMBUSTIBLE" in producto:
                        if es_pista_menor:
                            limite = 1000
                            tipo_limite = "1,000 L (Aceite - Pista Menor)"
                        else:
                            limite = 30280
                            tipo_limite = "30,280 L (Aceite - Pista Principal)"
                            
                    elif "MANCOL" in producto:
                        if es_pista_menor:
                            limite = 1000
                            tipo_limite = "1,000 L (Mancol - Pista Menor)"
                        else:
                            limite = 2500
                            tipo_limite = "2,500 L (Mancol - Pista Principal)"
                            
                    elif "ACONDICIONADOR" in producto or "NATURAMIN" in producto:
                        limite = 30
                        tipo_limite = "30 L/Kg (Aditivo de Alta Rotación)"
                    
                    # 4. Detonación de Alerta si se rompe el umbral
                    if saldo < limite:
                        alertas.append({
                            "📍 PISTA / ALMACÉN": pista,
                            "🧪 PRODUCTO QUÍMICO": producto,
                            "⚠️ SALDO ACTUAL": saldo,
                            "🛡️ LÍMITE DE SEGURIDAD": limite,
                            "📋 REGLA APLICADA": tipo_limite
                        })
                        
                # 5. Despliegue Visual
                if alertas:
                    st.error(f"🚨 **¡ALERTA ROJA! SE HAN DETECTADO {len(alertas)} INSUMOS POR DEBAJO DEL LÍMITE OPERATIVO:**")
                    df_alertas = pd.DataFrame(alertas).sort_values(by="📍 PISTA / ALMACÉN")
                    
                    # Formateo de números para lectura fácil
                    df_alertas["⚠️ SALDO ACTUAL"] = df_alertas["⚠️ SALDO ACTUAL"].apply(lambda x: f"{x:,.1f}".replace(",", "."))
                    df_alertas["🛡️ LÍMITE DE SEGURIDAD"] = df_alertas["🛡️ LÍMITE DE SEGURIDAD"].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                    
                    def pintar_rojo(val):
                        return ['background-color: #4a0000; color: white; font-weight: bold; border-bottom: 1px solid #ff4444;'] * len(val)
                        
                    st.dataframe(df_alertas.style.apply(pintar_rojo, axis=1), use_container_width=True, hide_index=True)
                else:
                    st.success("✅ **INVENTARIO ÓPTIMO:** Todos los insumos en todas las bodegas se encuentran por encima de los márgenes de seguridad. Operación asegurada.")
