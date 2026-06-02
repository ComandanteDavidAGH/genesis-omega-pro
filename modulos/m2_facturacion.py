import streamlit as st
import pandas as pd
import gspread
import io
import openpyxl
import re

def extraer_numero_local(val):
    try:
        v = str(val).replace(',', '.')
        v = re.sub(r'[^\d\.]', '', v)
        return float(v) if v else 0.0
    except: return 0.0

def ejecutar(extraer_numero):
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Facturación</h1>", unsafe_allow_html=True)
    
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
            with st.spinner("Desplegando Anclaje de Extracción Inteligente..."):
                try: 
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

                    lista_pistas = []
                    
                    for f in f_pistas:
                        nombre_archivo = f.name.lower()
                        bytes_f = io.BytesIO(f.getvalue())
                        dict_p = {}
                        
                        # 🎯 BLOQUEO DE HOJAS OCULTAS: data_only=True respeta la visibilidad del Excel
                        if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xlsm'):
                            wb_temp = openpyxl.load_workbook(bytes_f, data_only=True)
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
                            
                            idx_header = -1; col_finca = -1; col_pedido = -1; col_ha = -1
                            
                            for r in range(min(20, len(df))):
                                fila_textos = [str(x).strip().upper() for x in df.iloc[r].tolist()]
                                for c, val in enumerate(fila_textos):
                                    if any(palabra in val for palabra in ["FINCA", "HACIENDA", "CLIENTE"]): col_finca = c
                                    if any(palabra in val for palabra in ["PEDIDO", "ORDEN"]): col_pedido = c
                                    
                                    # 🎯 RADAR DE HECTÁREAS DE PISTA
                                    val_sin_esp = val.replace(" ", "")
                                    if ("HA" in val_sin_esp or "HECT" in val_sin_esp) and not "HORA" in val and not "H/H" in val and not "FECHA" in val:
                                        if "APLIC" in val or "GPS" in val or "FUMIG" in val: col_ha = c
                                        elif col_ha == -1: col_ha = c

                                if col_finca != -1: 
                                    idx_header = r
                                    break 
                                
                            if col_finca != -1:
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
                                                
                                    val_coctel = "S/N"
                                    for r_up in range(idx_header):
                                        fila_up = [str(x).strip().upper() for x in df.iloc[r_up].tolist()]
                                        for c_up, val in enumerate(fila_up):
                                            if "COCTEL" in val or "MEZCLA" in val:
                                                if c_up + 1 < len(fila_up) and fila_up[c_up+1] not in ["", "NAN", "NONE"]: 
                                                    val_coctel = str(df.iloc[r_up, c_up+1]).strip()
                                                elif c_up + 2 < len(fila_up) and fila_up[c_up+2] not in ["", "NAN", "NONE"]: 
                                                    val_coctel = str(df.iloc[r_up, c_up+2]).strip()
                                                
                                    # 🎯 EXTRACCIÓN DE HA DE PISTA
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
                        st.success(f"🛰️ Enlace Satelital Establecido. ¡Se extrajeron {len(lista_pistas)} misiones visibles con precisión! Pase al Módulo de Validación.")
                        st.balloons()
                    else:
                        st.session_state['df_pistas'] = pd.DataFrame()
                        st.warning("⚠️ La inteligencia no encontró misiones operativas en el documento de pista.")
                        
                except Exception as e: 
                    st.error(f"🚨 Error crítico en el escáner: {e}")
