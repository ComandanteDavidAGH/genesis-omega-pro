import streamlit as st
import pandas as pd
import gspread
import io
import openpyxl

def ejecutar(extraer_numero):
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Facturación</h1>", unsafe_allow_html=True)
    
    # --- 🛡️ BÚNKER DE MEMORIA ---
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
        st.markdown("### 🚁 3. Informes Pista")
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
            with st.spinner("Desplegando la Trampa de Diagnóstico..."):
                try: 
                    # 1. CARGA DE MATRICES BASE
                    nombre_sabana = f_sabana.name.lower()
                    if nombre_sabana.endswith(('.xlsx', '.xls')): st.session_state['df_sabana'] = pd.read_excel(f_sabana)
                    else:
                        try: st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='utf-8')
                        except:
                            f_sabana.seek(0)
                            st.session_state['df_sabana'] = pd.read_csv(f_sabana, sep=None, engine='python', encoding='latin1')
                    
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos) if f_pedidos.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(bytes_pedidos, sep=None, engine='python')
                        
                    if "gcp_credentials" in st.secrets: gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else: gc = gspread.service_account(filename='credenciales.json')
                    
                    url_boveda = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    boveda = gc.open_by_url(url_boveda)
                    
                    st.session_state['df_config'] = pd.DataFrame(boveda.worksheet("TABLA 2").get_all_values()[1:], columns=boveda.worksheet("TABLA 2").get_all_values()[0])
                    st.session_state['df_mezclas'] = pd.DataFrame(boveda.worksheet("DD_Mesclas").get_all_values()[1:], columns=boveda.worksheet("DD_Mesclas").get_all_values()[0])
                    st.session_state['df_config_base'] = pd.DataFrame(boveda.worksheet("Configuración").get_all_values()[1:], columns=boveda.worksheet("Configuración").get_all_values()[0])
                    
                    hoja_apoyo = boveda.worksheet("TABLA DE APOYO2023") 
                    datos_apoyo = hoja_apoyo.get_all_values()
                    
                    fila_titulos = 0
                    for i, fila in enumerate(datos_apoyo[:20]):
                        if any('FINCA' in str(celda).upper() for celda in fila): fila_titulos = i; break
                            
                    encabezados_crudos = datos_apoyo[fila_titulos]
                    encabezados_limpios = []
                    vientos = {}
                    for col in encabezados_crudos:
                        col_str = str(col).strip()
                        if col_str == "": col_str = "Vacio"
                        if col_str in vientos:
                            vientos[col_str] += 1
                            encabezados_limpios.append(f"{col_str}_{vientos[col_str]}")
                        else:
                            vientos[col_str] = 0
                            encabezados_limpios.append(col_str)
                            
                    st.session_state['df_apoyo'] = pd.DataFrame(datos_apoyo[fila_titulos+1:], columns=encabezados_limpios)

                    # 2. 🔥 TRAMPA EN EL ESCÁNER DE PISTAS
                    lista_pistas = []
                    
                    # 💡 EXPANDER DE DIAGNÓSTICO VISUAL
                    with st.expander("🕵️ RADAR DE DIAGNÓSTICO ACTIVO (ABRIR PARA VER RADIOGRAFÍA)", expanded=True):
                        st.warning("Analizando el archivo de pista tal como lo ve la máquina...")
                        
                        for f in f_pistas:
                            nombre_archivo = f.name.lower()
                            bytes_f = io.BytesIO(f.getvalue())
                            dict_p = {}
                            
                            if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xlsm'):
                                wb_temp = openpyxl.load_workbook(bytes_f, read_only=True)
                                hojas_visibles = [ws.title for ws in wb_temp.worksheets if ws.sheet_state == 'visible']
                                bytes_f.seek(0)
                                if hojas_visibles: dict_p = pd.read_excel(bytes_f, sheet_name=hojas_visibles, header=None)
                            elif nombre_archivo.endswith('.xls'): dict_p = pd.read_excel(bytes_f, sheet_name=None, header=None)
                            else:
                                try: dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', header=None)}
                                except:
                                    bytes_f.seek(0)
                                    dict_p = {"Datos_CSV": pd.read_csv(bytes_f, sep=None, engine='python', encoding='latin1', header=None)}
                                
                            for n, df in dict_p.items():
                                df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
                                
                                # IMPRIMIR LA RADIOGRAFÍA EN PANTALLA
                                st.write(f"**Archivo:** {f.name} | **Pestaña:** {n}")
                                st.dataframe(df.head(20)) # Muestra las primeras 20 filas crudas
                                
                                col_finca = -1
                                coctel_actual = "S/N"
                                
                                for r in range(len(df)):
                                    fila_textos = [str(x).strip().upper() for x in df.iloc[r].tolist()]
                                    
                                    # Rastreo muy suave de finca
                                    for c_idx, celda in enumerate(fila_textos):
                                        if any(palabra in celda for palabra in ["FINCA", "HACIENDA", "CLIENTE"]):
                                            col_finca = c_idx
                                            break
                                            
                                    if col_finca != -1:
                                        val_finca = str(df.iloc[r, col_finca]).strip()
                                        if any(palabra in val_finca.upper() for palabra in ["FINCA", "HACIENDA", "CLIENTE", "TOTAL", "NAN", "NONE"]): continue
                                        if val_finca == "": continue
                                            
                                        datos_fila = df.iloc[r].to_dict()
                                        lista_pistas.append({
                                            "ORIGEN": f"{f.name} | Fila {r+1}", 
                                            "COCTEL": "INTENTO CAPTURA", 
                                            "FINCA_INFORME": val_finca, 
                                            "PEDIDO_SAP": "S/N",
                                            "DATOS_FILA": datos_fila
                                        })

                    # 🎯 INYECCIÓN DEL SEÑUELO SI TODO FALLA
                    if not lista_pistas:
                        st.error("🚨 El escáner no pudo encontrar NINGUNA finca real en el documento.")
                        st.info("🎣 Inyectando SEÑUELO FANTASMA para probar la conexión con el Módulo 3...")
                        lista_pistas.append({
                            "ORIGEN": "SEÑUELO DE PRUEBA | IGNORAR",
                            "COCTEL": "MEZCLA FANTASMA",
                            "FINCA_INFORME": "FINCA FANTASMA DE PRUEBA",
                            "PEDIDO_SAP": "1234567",
                            "DATOS_FILA": {0: "Dato 1", 1: "Dato 2"}
                        })

                    st.session_state['df_pistas'] = pd.DataFrame(lista_pistas)
                    st.success(f"🛰️ Enlace Satelital Establecido. Pase al Módulo de Validación.")
                    st.balloons()
                        
                except Exception as e: 
                    st.error(f"🚨 Error crítico en el escáner: {e}")
