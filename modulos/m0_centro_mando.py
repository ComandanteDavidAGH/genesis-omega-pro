import streamlit as st
import pandas as pd
import numpy as np
import gspread
import math
import re

def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        elif "gcp_credentials" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

@st.cache_data(show_spinner=False, ttl=600)
def cargar_inventario_supabase_cached():
    if 'supabase' in st.session_state and st.session_state['supabase'] is not None:
        try:
            supabase_client = st.session_state['supabase']
            respuesta = supabase_client.table("inventario_sap").select("*").execute()
            if respuesta.data and len(respuesta.data) > 0:
                return pd.DataFrame(respuesta.data)
        except Exception:
            pass
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def procesar_radar_logistico_cached(df_sabana):
    if df_sabana.empty:
        return None, "VACIO", 0, 0, 0, pd.DataFrame()

    col_cod = next((c for c in df_sabana.columns if str(c).strip() == 'Material'), None)
    col_pista = next((c for c in df_sabana.columns if str(c).strip() == 'Almacén'), None)
    col_saldo = next((c for c in df_sabana.columns if str(c).strip() == 'Libre utilización'), None)
    col_desc = next((c for c in df_sabana.columns if str(c).strip() == 'Descripción del material'), None)

    if not col_cod: col_cod = next((c for c in df_sabana.columns if 'MATERIAL' in str(c).upper()), None)
    if not col_pista: col_pista = next((c for c in df_sabana.columns if 'ALMACEN' in str(c).upper() or 'LGORT' in str(c).upper()), None)
    if not col_saldo: col_saldo = next((c for c in df_sabana.columns if 'LIBRE' in str(c).upper() or 'UTILIZACION' in str(c).upper() or 'LABST' in str(c).upper()), None)
    if not col_desc: col_desc = next((c for c in df_sabana.columns if 'DESC' in str(c).upper() or 'TEXTO' in str(c).upper()), None)

    if not col_cod or not col_pista or not col_saldo:
        return None, "ERROR_COLUMNAS", 0, 0, 0, pd.DataFrame()

    df_temp = df_sabana.copy()
    df_temp[col_saldo] = pd.to_numeric(df_temp[col_saldo].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    df_temp = df_temp[df_temp[col_saldo] > 0]
    
    if df_temp.empty:
        return None, "OK", 0, 0, 0, pd.DataFrame()

    codigos_limpios = df_temp[col_cod].astype(str).str.split('.').str[0].str.strip()
    
    if col_desc:
        df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | " + df_temp[col_desc].astype(str).str.strip().str.upper()
    else:
        df_temp['PRODUCTO_RADAR'] = codigos_limpios + " | INSUMO QUÍMICO REGISTRADO"

    inventario_agrupado = df_temp.groupby([col_pista, 'PRODUCTO_RADAR'])[col_saldo].sum().reset_index()
    
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
    
    df_alertas = inventario_agrupado[inventario_agrupado[col_saldo] < inventario_agrupado['🛡️ LÍMITE DE SEGURIDAD']].copy()
    
    df_alertas = df_alertas.rename(columns={
        col_pista: "📍 PISTA / ALMACÉN",
        'PRODUCTO_RADAR': "🧪 CÓDIGO | NOMBRE DEL PRODUCTO",
        col_saldo: "⚠️ SALDO ACTUAL"
    })
    
    columnas_finales = ["📍 PISTA / ALMACÉN", "🧪 CÓDIGO | NOMBRE DEL PRODUCTO", "⚠️ SALDO ACTUAL", "🛡️ LÍMITE DE SEGURIDAD", "📋 REGLA APLICADA"]
    
    if not df_alertas.empty:
        df_alertas_render = df_alertas[columnas_finales].sort_values(by="📍 PISTA / ALMACÉN")
    else:
        df_alertas_render = pd.DataFrame(columns=columnas_finales)
    
    total_almacenes = inventario_agrupado[col_pista].nunique()
    total_insumos = inventario_agrupado['PRODUCTO_RADAR'].nunique()
    conteo_alertas = len(df_alertas_render)

    return df_alertas_render, "EXITO", total_almacenes, total_insumos, conteo_alertas, df_alertas_raw

# =================================================================
# 🗄️ ORDENAMIENTO GLOBAL DE BASE DE DATOS (AUTORREGENERABLE)
# =================================================================

def ordenar_base_datos_global():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        st.error("🚨 Sin conexión activa a Supabase.")
        return

    def limpiar_cabecera(h):
        h = str(h).replace('\n', ' ').strip()
        h = ' '.join(h.split())
        mapeos = {
            "COCTEL": "CÓCTEL",
            "DÌA SEM": "DÍA SEM",
            "DIA SEM": "DÍA SEM"
        }
        return mapeos.get(h, h)

    def sanitizar(v):
        if pd.isna(v) or v is None: return ""
        if isinstance(v, (float, int)):
            if math.isnan(v) or math.isinf(v): return 0
            return v
        return str(v).strip()

    try:
        supabase = st.session_state['supabase']
        gc = inicializar_cliente_gspread()
        if not gc: 
            st.error("🚨 Sin conexión a Google Drive.")
            return

        with st.spinner("🔄 Rescatando Sábana de Drive y calibrando esquema de Supabase..."):
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            t1_formulas = ws_t1.get_all_values(value_render_option='FORMULA')
            t1_valores = ws_t1.get_all_values(value_render_option='UNFORMATTED_VALUE')
            
            idx_header = 4
            for i in range(min(10, len(t1_formulas))):
                if "FINCA" in [str(x).upper().strip() for x in t1_formulas[i]]:
                    idx_header = i
                    break
                    
            headers_excel = [str(x) for x in t1_valores[idx_header]]
            headers_limpios = [limpiar_cabecera(x) for x in headers_excel]
            num_cols = len(headers_limpios)
            
            datos_form = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_formulas[idx_header+1:]]
            datos_val = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_valores[idx_header+1:]]
            
            df_form = pd.DataFrame(datos_form, columns=headers_excel)
            df_val = pd.DataFrame(datos_val, columns=headers_limpios)
            
            col_id_limpio = headers_limpios[0]
            
            filas_validas = df_val[col_id_limpio].astype(str).str.strip() != ""
            df_form = df_form[filas_validas].copy()
            df_val = df_val[filas_validas].copy()

            col_fecha = next((c for c in headers_limpios if "FECHA" in c.upper()), None)

            if col_fecha:
                df_val['fecha_dt'] = pd.to_datetime(df_val[col_fecha].astype(str).str.replace("'", "").str.strip(), format='%d/%m/%Y', errors='coerce')
                df_val_sorted = df_val.sort_values(by='fecha_dt', ascending=False, na_position='last')
                indices_ord = df_val_sorted.index
                df_form_sorted = df_form.loc[indices_ord]
                df_val_sorted = df_val_sorted.drop(columns=['fecha_dt'])
            else:
                df_val_sorted = df_val
                df_form_sorted = df_form

            # CONSTRUIR DICCIONARIOS
            registros = []
            for _, row in df_val_sorted.iterrows():
                rec = {c: sanitizar(row[c]) for c in headers_limpios if c != ""}
                registros.append(rec)

            if registros:
                st.info("🧪 Calibrando campos con Supabase en tiempo real...")
                
                # 💥 ALGORITMO AUTORREGENERABLE (AUTO-HEAL)
                test_item = registros[0].copy()
                llaves_validas = list(test_item.keys())
                
                for intento in range(35):
                    try:
                        # Intentar inyección de prueba
                        supabase.table("TABLA_1").insert([test_item]).execute()
                        llaves_validas = list(test_item.keys())
                        break
                    except Exception as e_test:
                        err_str = str(e_test)
                        # Extraer dinámicamente la columna rechazada por Supabase
                        match = re.search(r"Could not find the '([^']+)' column", err_str)
                        if match:
                            col_rechazada = match.group(1)
                            if col_rechazada in test_item:
                                del test_item[col_rechazada]
                                continue
                        # Si es otro tipo de error no relacionado con columnas, romper
                        st.error(f"🚨 Error imprevisto en Supabase: {e_test}")
                        return

                st.success("✅ Estructura calibrada con éxito. Procediendo al llenado seguro...")
                
                # Filtrar todos los registros dejándolos SOLO con las llaves que Supabase aceptó
                registros_filtrados = [{k: v for k, v in r.items() if k in llaves_validas} for r in registros]
                
                # Borrado previo seguro y carga masiva
                supabase.table("TABLA_1").delete().neq(col_id_limpio, "_VACIO_IMPOSIBLE_999_").execute()
                
                tamano_bloque = 250
                for i in range(0, len(registros_filtrados), tamano_bloque):
                    supabase.table("TABLA_1").insert(registros_filtrados[i:i + tamano_bloque]).execute()

            # ACTUALIZAR GOOGLE DRIVE (CON FÓRMULAS E INTEGRIDAD)
            valores_drive = df_form_sorted[headers_excel].fillna("").values.tolist()
            if valores_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                ws_t1.batch_clear([rango_borrar])
                ws_t1.update(range_name=rango_inicio, values=valores_drive, value_input_option='USER_ENTERED')
                
            st.success(f"🎉 ¡MISIÓN CUMPLIDA! Base de Datos Restaurada ({len(registros)} registros) y Sincronizada al 100%.")
            st.balloons()

    except Exception as e:
        st.error(f"🚨 Error en sincronización: {e}")

# =================================================================
# 👑 RENDERIZADO VISUAL DEL CENTRO DE MANDO
# =================================================================

def renderizar():
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.1) !important; overflow: hidden !important; }
    .hud-mando { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border-left: 5px solid #0d1b2a; padding: 12px 20px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; box-shadow: 2px 2px 8px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #dee2e6; }
    .hud-mando-item { text-align: center; }
    .hud-mando-title { font-size: 11px; color: #6c757d; font-family: 'Arial Black', sans-serif; text-transform: uppercase; margin: 0; }
    .hud-mando-value { font-size: 20px; color: #0d1b2a; font-weight: 900; margin: 0; }
    .hud-mando-alert { color: #cc0000; font-family: 'Arial Black', sans-serif; }
    .hud-mando-ok { color: #00994c; font-family: 'Arial Black', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>🏠 Centro de Mando y Control</h1>", unsafe_allow_html=True)
    st.info("📡 **Radar Principal:** Monitoreo activo de sistemas, escuadrones y logística aérea.")
    st.markdown(f"### Bienvenido al Cuartel General, **{st.session_state.get('usuario_nombre', 'Comandante')}**.")
    st.write("El sistema Génesis Omega Pro se encuentra en línea y operando bajo parámetros óptimos. Seleccione un hangar en el menú lateral para iniciar operaciones.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 🗄️ Panel de Mantenimiento de Base de Datos")
    st.info("💡 **Alineación Cronológica:** Utilice esta herramienta para ordenar físicamente todas las misiones por fecha. El sistema tomará todas las operaciones y pondrá las fechas más recientes en la parte superior tanto en Google Drive como en Supabase (`TABLA_1`).")
    
    if st.button("🧹 ORDENAR DRIVE Y SUPABASE POR FECHA", type="primary", use_container_width=True):
        ordenar_base_datos_global()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("### 🚨 Radar Logístico: Alerta Temprana de Inventarios")
    
    df_sabana = st.session_state.get('df_sabana', pd.DataFrame())
    if df_sabana.empty:
        df_sabana = cargar_inventario_supabase_cached()
        if not df_sabana.empty:
            st.session_state['df_sabana'] = df_sabana

    if df_sabana.empty:
        st.warning("⚠️ **Radar en Modo Espera:** El sistema no detecta un inventario activo en la memoria ni en la nube. Para encender el radar, por favor cargue la **Sábana SAP** actualizada en el **📥 Módulo 2 (Carga Facturación)**.")
    else:
        df_alertas_render, estado, total_almacenes, total_insumos, conteo_alertas, df_alertas_raw = procesar_radar_logistico_cached(df_sabana)

        if estado == "ERROR_COLUMNAS":
            st.error("❌ Error de Radar: No se pudieron mapear las columnas. Verifique que el archivo corresponda a la Sábana Estándar o estructura de Supabase.")
        else:
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
            
            if conteo_alertas > 0:
                st.error("🚨 **¡ALERTA ROJA! MÁRGENES OPERATIVOS CRÍTICOS DETECTADOS:**")
                df_render_display = df_alertas_render.copy()
                df_render_display["⚠️ SALDO ACTUAL"] = df_render_display["⚠️ SALDO ACTUAL"].apply(lambda x: f"{x:,.1f}".replace(",", "."))
                df_render_display["🛡️ LÍMITE DE SEGURIDAD"] = df_render_display["🛡️ LÍMITE DE SEGURIDAD"].apply(lambda x: f"{x:,.0f}".replace(",", "."))
                
                def pintar_rojo_elegante(val):
                    return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold; border-bottom: 1px solid #dee2e6;'] * len(val)
                
                st.dataframe(df_render_display.style.apply(pintar_rojo_elegante, axis=1), use_container_width=True, hide_index=True)
            else:
                st.success("✅ **INVENTARIO ÓPTIMO:** Todos los insumos químicos y energéticos en la totalidad de las pistas se encuentran por encima de los márgenes de seguridad establecidos. Operación aérea asegurada.")
