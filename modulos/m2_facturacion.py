import streamlit as st
import pandas as pd
import io
import re
import json
import concurrent.futures

# =================================================================
# ⚡ MOTORES DE CARGA Y DESCARGA EN CACHÉ ULTRARRÁPIDA
# =================================================================
URL_BOVEDA_MAESTRA = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"

def extraer_numero_local(val):
    try:
        v = str(val).replace(',', '.')
        v = re.sub(r'[^\d\.]', '', v)
        return float(v) if v else 0.0
    except Exception:
        return 0.0

@st.cache_data(show_spinner=False, ttl=3600)
def cargar_tablas_maestras_m2_cached():
    """
    Intenta recuperar las 4 tablas maestras desde Supabase primero (milisegundos).
    Si no existen, las descarga de Google Sheets en paralelo.
    """
    resultados = {}
    mapeo_tablas = {
        'df_config': ('config_tabla2', "TABLA%202"),
        'df_mezclas': ('dd_mezclas', "DD_Mesclas"),
        'df_config_base': ('configuracion_base', "Configuraci%C3%B3n"),
        'df_apoyo_raw': ('tabla_apoyo_raw', "TABLA%20DE%20APOYO2023")
    }

    # 1. Intentar vía Supabase (Velocidad Luz)
    if 'supabase' in st.session_state and st.session_state['supabase'] is not None:
        try:
            supabase_client = st.session_state['supabase']
            for key, (table_name, _) in mapeo_tablas.items():
                resp = supabase_client.table(table_name).select("*").execute()
                if resp.data and len(resp.data) > 0:
                    resultados[key] = pd.DataFrame(resp.data)
        except Exception:
            resultados = {}

    # 2. Fallback a Google Sheets solo si faltan tablas
    if len(resultados) < 4:
        url_base = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/gviz/tq?tqx=out:csv&sheet="
        urls_a_bajar = {k: f"{url_base}{v[1]}" for k, v in mapeo_tablas.items() if k not in resultados}

        def fetch_url(key, url):
            try:
                return key, pd.read_csv(url, skiprows=(0 if key == 'df_apoyo_raw' else 1))
            except Exception:
                return key, pd.DataFrame()

        try:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                res_fetched = dict(executor.map(lambda item: fetch_url(*item), urls_a_bajar.items()))
                resultados.update(res_fetched)
        except Exception:
            pass

    return resultados

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE FACTURACIÓN
# =================================================================

def ejecutar(extraer_numero):
    st.markdown("""
    <style>
    .titulo-principal { 
        color: #0d1b2a; 
        border-bottom: 3px solid #d4af37; 
        padding-bottom: 5px; 
        font-family: 'Arial Black', sans-serif; 
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>📥 Zona de Aterrizaje Facturación</h1>", unsafe_allow_html=True)
    
    if 'mem_sabana' not in st.session_state: st.session_state['mem_sabana'] = None
    if 'name_sabana' not in st.session_state: st.session_state['name_sabana'] = None
    if 'mem_pedidos' not in st.session_state: st.session_state['mem_pedidos'] = None
    if 'name_pedidos' not in st.session_state: st.session_state['name_pedidos'] = None
    
    if 'df_pistas' not in st.session_state: st.session_state['df_pistas'] = pd.DataFrame()
    if 'df_apoyo' not in st.session_state: st.session_state['df_apoyo'] = pd.DataFrame()

    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        if st.session_state['mem_sabana'] is None:
            f_sabana_up = st.file_uploader("Inventario, Precios y Lotes", type=["xlsx", "xls", "csv"], key="sab")
            if f_sabana_up:
                st.session_state['mem_sabana'] = f_sabana_up.getvalue()
                st.session_state['name_sabana'] = f_sabana_up.name
                st.rerun()
        else:
            st.success(f"✅ Sábana en memoria: {st.session_state['name_sabana']}")
            if st.button("🔄 Cambiar Sábana", use_container_width=True):
                st.session_state['mem_sabana'] = None
                st.rerun()

    with c2:
        st.markdown("### 📝 2. Pedidos SAP")
        if st.session_state['mem_pedidos'] is None:
            f_pedidos_up = st.file_uploader("Planificación (Finca/Cantidades)", type=["xlsx", "xls", "csv"], key="ped")
            if f_pedidos_up:
                st.session_state['mem_pedidos'] = f_pedidos_up.getvalue()
                st.session_state['name_pedidos'] = f_pedidos_up.name
                st.rerun()
        else:
            st.success(f"✅ Pedidos en memoria: {st.session_state['name_pedidos']}")
            if st.button("🔄 Cambiar Pedidos", use_container_width=True):
                st.session_state['mem_pedidos'] = None
                st.rerun()

    with c3:
        st.markdown("### 🚀 3. Informes Pista")
        f_pistas = st.file_uploader("Reportes Reales", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="pis")

    f_sabana = None
    if st.session_state['mem_sabana']:
        f_sabana = io.BytesIO(st.session_state['mem_sabana'])
        f_sabana.name = st.session_state['name_sabana']
        
    f_pedidos = None
    if st.session_state['mem_pedidos']:
        f_pedidos = io.BytesIO(st.session_state['mem_pedidos'])
        f_pedidos.name = st.session_state['name_pedidos']

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas:
            with st.spinner("Desplegando Anclaje de Extracción Inteligente (Optimizado)..."):
                try: 
                    # 1. Carga directa de Sábana y Pedidos
                    nombre_sabana = f_sabana.name.lower()
                    if nombre_sabana.endswith(('.xlsx', '.xls')): 
                        st.session_state['df_sabana'] = pd.read_excel(f_sabana)
                    else:
                        try: 
                            st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='utf-8')
                        except Exception:
                            f_sabana.seek(0)
                            st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='latin1')
                    
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos) if f_pedidos.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    # ⚡ 2. Carga en caché ultrarrápida de tablas maestras
                    resultados = cargar_tablas_maestras_m2_cached()

                    # Asignación a memoria de sesión
                    st.session_state['df_config'] = resultados.get('df_config', pd.DataFrame())
                    st.session_state['df_mezclas'] = resultados.get('df_mezclas', pd.DataFrame())
                    st.session_state['df_config_base'] = resultados.get('df_config_base', pd.DataFrame())
                    df_apoyo_raw = resultados.get('df_apoyo_raw', pd.DataFrame())
                    
                    # Limpieza veloz de TABLA DE APOYO
                    df_apoyo_final = pd.DataFrame()
                    if not df_apoyo_raw.empty:
                        fila_titulos = 0
                        for i in range(min(20, len(df_apoyo_raw))):
                            if df_apoyo_raw.iloc[i].astype(str).str.upper().str.contains('FINCA').any():
                                fila_titulos = i
                                break
                                
                        encabezados_crudos = df_apoyo_raw.iloc[fila_titulos].tolist()
                        encabezados_limpios = []
                        vientos = {}
                        for col in encabezados_crudos:
                            col_str = str(col).strip()
                            if col_str in ["", "nan", "NaN", "None"]: col_str = "Vacio"
                            if col_str in vientos:
                                vientos[col_str] += 1
                                encabezados_limpios.append(f"{col_str}_{vientos[col_str]}")
                            else:
                                vientos[col_str] = 0
                                encabezados_limpios.append(col_str)
                                
                        df_apoyo_final = df_apoyo_raw.iloc[fila_titulos+1:].copy()
                        df_apoyo_final.columns = encabezados_limpios
                    st.session_state['df_apoyo'] = df_apoyo_final

                    # ⚡ 3. Procesamiento optimizado de Pistas mediante pd.ExcelFile (Una sola lectura)
                    lista_pistas = []
                    
                    for f in f_pistas:
                        nombre_archivo = f.name.lower()
                        bytes_f = io.BytesIO(f.getvalue())
                        dict_p = {}
                        
                        if nombre_archivo.endswith(('.xlsx', '.xlsm', '.xls')):
                            xls_file = pd.ExcelFile(bytes_f)
                            dict_p = {sheet: xls_file.parse(sheet, header=None) for sheet in xls_file.sheet_names}
                        else:
                            try:
                                dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', header=None)}
                            except Exception:
                                bytes_f.seek(0)
                                dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', encoding='latin1', header=None)}
                            
                        for n, df in dict_p.items():
                            if df.empty: continue
                            df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                            
                            idx_header = -1; col_finca = -1; col_pedido = -1; col_ha = -1
                            
                            for r in range(min(20, len(df))):
                                fila_textos = [str(x).strip().upper() for x in df.iloc[r].tolist()]
                                for c, val in enumerate(fila_textos):
                                    if any(palabra in val for palabra in ["FINCA", "HACIENDA", "CLIENTE"]): col_finca = c
                                    if any(palabra in val for palabra in ["PEDIDO", "ORDEN"]): col_pedido = c
                                    
                                    val_sin_esp = val.replace(" ", "")
                                    if ("HA" in val_sin_esp or "HECT" in val_sin_esp) and not "HORA" in val and not "H/H" in val and not "FECHA" in val:
                                        if "APLIC" in val or "GPS" in val or "FUMIG" in val: col_ha = c
                                        elif col_ha == -1: col_ha = c

                                if col_finca != -1: 
                                    idx_header = r
                                    break 
                                
                            if col_finca != -1:
                                val_coctel_global = "S/N"
                                for r_up in range(idx_header):
                                    fila_up = [str(x).strip().upper() for x in df.iloc[r_up].tolist()]
                                    for c_up, val in enumerate(fila_up):
                                        if "COCTEL" in val or "MEZCLA" in val:
                                            if c_up + 1 < len(fila_up) and fila_up[c_up+1] not in ["", "NAN", "NONE"]: 
                                                val_coctel_global = str(df.iloc[r_up, c_up+1]).strip()
                                            elif c_up + 2 < len(fila_up) and fila_up[c_up+2] not in ["", "NAN", "NONE"]: 
                                                val_coctel_global = str(df.iloc[r_up, c_up+2]).strip()
                                                
                                for r in range(idx_header + 1, len(df)):
                                    val_finca = str(df.iloc[r, col_finca]).strip()
                                    if val_finca.upper() in ["", "NAN", "NONE", "TOTAL"] or "TOTAL" in val_finca.upper(): continue
                                    if len(val_finca) < 2: continue 
                                    
                                    fila_actual_textos = [str(x).strip().upper() for x in df.iloc[r].tolist()]
                                    
                                    val_pedido = "S/N"
                                    if col_pedido != -1 and col_pedido < len(df.columns):
                                        v_p = str(df.iloc[r, col_pedido]).split('.')[0].strip()
                                        if v_p.isdigit() and len(v_p) >= 6: val_pedido = v_p
                                            
                                    if val_pedido == "S/N":
                                        for celda in reversed(fila_actual_textos):
                                            c_clean = celda.split('.')[0].strip()
                                            if c_clean.isdigit() and len(c_clean) >= 6:
                                                val_pedido = c_clean; break
                                                
                                    val_coctel = val_coctel_global
                                    
                                    val_ha_pista = 0.0
                                    if col_ha != -1 and col_ha < len(df.columns):
                                        val_ha_pista = extraer_numero_local(df.iloc[r, col_ha])
                                            
                                    lista_pistas.append({
                                        "ORIGEN": f"{f.name} | {n}", 
                                        "COCTEL": val_coctel, 
                                        "FINCA_INFORME": val_finca, 
                                        "PEDIDO_SAP": val_pedido,
                                        "HA_PISTA": val_ha_pista,
                                        "DATOS_FILA": df.iloc[r].to_dict()
                                    })

                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                        
                        # =========================================================
                        # 🛡️ MEJORA: MEMORIA PURA PARA MÓDULO 3
                        # Las misiones quedan cargadas directamente en RAM sin 
                        # depender de tablas inexistentes en Supabase.
                        # =========================================================
                                
                        st.success(f"🛰️ Enlace Satelital Establecido. ¡Se extrajeron {len(lista_pistas)} misiones visibles con precisión! Pase al Módulo de Validación.")
                        st.balloons()
                    else:
                        st.session_state['df_pistas'] = pd.DataFrame()
                        st.warning("⚠️ La inteligencia no encontró misiones operativas en los documentos de pista cargados.")
                        
                except Exception as e: 
                    st.error(f"🚨 Error crítico en el escáner del procesamiento maestro: {e}")
        else:
            st.warning("⚠️ Asegúrese de haber subido la Sábana SAP, el Pedido SAP y al menos un Informe de Pista antes de iniciar el procesamiento.")

if __name__ == "__main__":
    pass
