import streamlit as st
import pandas as pd
import gspread

def ejecutar(extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro):
    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    
    # --- 🧮 NUEVA SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Copia Rápida)")
        st.info("💡 Obtenga la lista de precios exactos multiplicados por el margen de cada perfil, listos para copiar y pegar en SAP.")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Conectando con la Bóveda de Configuración..."):
                try:
                    if "gcp_credentials" in st.secrets:
                        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
                    else:
                        gc = gspread.service_account(filename='credenciales.json')
                        
                    sh_gen = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
                    raw_config = sh_gen.worksheet("Configuración").get_all_values()
                    
                    lista_precios = []
                    for row in raw_config:
                        if len(row) > 9:
                            prod = str(row[8]).upper().strip()
                            
                            # 🛡️ FILTRO ANTI-FANTASMAS (Destruye 0, 0.0, 0.00 y vacíos)
                            es_cero_basura = False
                            try:
                                if float(prod) == 0: es_cero_basura = True
                            except ValueError:
                                pass
                                
                            if prod and prod != "PRODUCTO" and "INVENTARIO" not in prod and not es_cero_basura:
                                costo_base = extraer_numero(row[9])
                                if costo_base > 0:
                                    lista_precios.append({
                                        "PRODUCTO": prod,
                                        "COSTO BASE": costo_base,
                                        "TERCERO (+45.1%)": round(costo_base * 1.451, 0),
                                        "AFILIADO (+16.4%)": round(costo_base * 1.164, 0),
                                        "COOPERATIVA / SOCIO (+11.2%)": round(costo_base * 1.112, 0),
                                        "ORGÁNICO (+1.1%)": round(costo_base * 1.011, 0)
                                    })
                    
                    if lista_precios:
                        df_tarifario = pd.DataFrame(lista_precios).sort_values(by="PRODUCTO").reset_index(drop=True)
                        st.session_state['df_tarifario'] = df_tarifario
                        st.success(f"✅ Tarifario cargado: {len(lista_precios)} productos ordenados alfabéticamente (A-Z).")
                    else:
                        st.warning("⚠️ El escáner no encontró productos con precios válidos.")
                except Exception as e:
                    st.error(f"🚨 Error al generar tarifario: {e}")
                    
        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_t = st.session_state['df_tarifario']
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Copia Individual (Por Producto)"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes (Ordenada por Producto)")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### Caja de Copiado Masivo (Formación Alineada)")
                col_margen = st.selectbox("1️⃣ Seleccione el Perfil de Productor:", 
                                          ["TERCERO (+45.1%)", "AFILIADO (+16.4%)", "COOPERATIVA / SOCIO (+11.2%)", "ORGÁNICO (+1.1%)", "COSTO BASE"])
                
                incluir_nombres = st.toggle("🔘 Incluir Nombre del Producto (Alineación Perfecta)", value=False)
                st.caption(f"2️⃣ Copie la lista haciendo clic en el ícono de la esquina superior derecha:")
                
                if col_margen in df_t.columns:
                    if incluir_nombres:
                        max_len = df_t["PRODUCTO"].apply(len).max() + 4
                        
                        lista_textos = []
                        for _, row in df_t.iterrows():
                            nombre = str(row["PRODUCTO"]).strip()
                            precio = fmt_sap(row[col_margen])
                            nombre_alineado = nombre.ljust(max_len)
                            lista_textos.append(f"{nombre_alineado}\t{precio}")
                        texto_para_copiar = "\n".join(lista_textos)
                    else:
                        lista_textos = [fmt_sap(x) for x in df_t[col_margen]]
                        texto_para_copiar = "\n".join(lista_textos)
                        
                    st.code(texto_para_copiar, language="text")
                    
            with t3:
                st.markdown("#### Búsqueda y Copia Rápida Individual (Modo Francotirador)")
                prod_sel = st.selectbox("🔍 Buscar Producto Específico:", df_t["PRODUCTO"].tolist())
                
                if prod_sel:
                    datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
                    st.info(f"🎯 Valores calculados para: **{prod_sel}**")
                    
                    c1, c2, c3, c4, c5 = st.columns(5)
                    
                    with c1:
                        st.caption("Costo Base")
                        st.code(fmt_sap(datos_prod["COSTO BASE"]), language="text")
                    with c2:
                        st.caption("Orgánico")
                        st.code(fmt_sap(datos_prod["ORGÁNICO (+1.1%)"]), language="text")
                    with c3:
                        st.caption("Socio / Coop")
                        st.code(fmt_sap(datos_prod["COOPERATIVA / SOCIO (+11.2%)"]), language="text")
                    with c4:
                        st.caption("Afiliado")
                        st.code(fmt_sap(datos_prod["AFILIADO (+16.4%)"]), language="text")
                    with c5:
                        st.caption("Tercero")
                        st.code(fmt_sap(datos_prod["TERCERO (+45.1%)"]), language="text")
                        
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    semana_target = st.select_slider("Semana a actualizar:", options=list(range(1, 53)), value=19)

    if st.button("🚀 EJECUTAR OMEGA V12", use_container_width=True):
        try:
            with st.spinner(f"Sincronizando Semana {semana_target} al estilo Macro..."):
                if "gcp_credentials" in st.secrets:
                    cred_dict = dict(st.secrets["gcp_credentials"])
                    gc = gspread.service_account_from_dict(cred_dict)
                else:
                    gc = gspread.service_account(filename='credenciales.json')

                url_gen = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                sh_gen = gc.open_by_url(url_gen)
                
                raw_config = sh_gen.worksheet("Configuración").get_all_values(value_render_option='UNFORMATTED_VALUE')
                dict_precios = {}
                for row in raw_config:
                    if len(row) > 9:
                        prod = limpiar_texto_vba(row[8])
                        if prod and prod != "PRODUCTO":
                            dict_precios[prod] = val_seguro(row[9])

                raw_mezclas = sh_gen.worksheet("DD_Mesclas").get_all_values(value_render_option='UNFORMATTED_VALUE')
                dict_dosis = {}
                for row in raw_mezclas[12:]: 
                    if len(row) > 10:
                        prod_m = limpiar_texto_vba(row[9])
                        if prod_m:
                            dict_dosis[prod_m] = val_seguro(row[10])

                url_dest = "https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit"
                sh_dest = gc.open_by_url(url_dest)
                ws_datos = sh_dest.worksheet("DATOS")
                datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                
                col_semana = -1
                for i, v in enumerate(datos_dest[6]):
                    if str(v).strip() == str(semana_target):
                        col_semana = i + 1
                        break
                
                if col_semana == -1:
                    st.error(f"❌ No se halló la semana {semana_target} en la Fila 7.")
                else:
                    updates = []
                    for r_idx, row in enumerate(datos_dest):
                        n_fila = r_idx + 1
                        if n_fila < 8 or len(row) < 4: continue
                        
                        tipo_tabla = limpiar_texto_vba(row[1]) 
                        producto_dest = limpiar_texto_vba(row[3])
                        
                        if not producto_dest: continue
                        
                        if producto_dest in dict_precios:
                            precio_unitario = dict_precios[producto_dest]
                            if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                if producto_dest in dict_dosis:
                                    dosis_valor = dict_dosis[producto_dest]
                                    valor_final = precio_unitario * dosis_valor
                                else:
                                    valor_final = precio_unitario
                            else:
                                valor_final = precio_unitario
                                
                            updates.append({
                                'range': gspread.utils.rowcol_to_a1(n_fila, col_semana),
                                'values': [[valor_final]]
                            })

                    if updates:
                        ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                        st.success(f"🎯 IMPACTO PERFECTO. {len(updates)} precios inyectados con precisión absoluta.")
                        st.balloons()
                    else:
                        st.warning("⚠️ No se encontraron productos coincidentes.")

        except Exception as e:
            st.error(f"🚨 FALLA DEL SISTEMA: {e}")
