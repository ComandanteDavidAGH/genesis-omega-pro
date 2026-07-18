import streamlit as st
import pandas as pd
import gspread

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

def purificar_y_convertir_precio(valor_crudo):
    if not valor_crudo:
        return 0.0
    val_str = str(valor_crudo).replace("$", "").replace("COP", "").replace(" ", "").strip()
    if "," in val_str and "." in val_str:
        if val_str.find(".") < val_str.find(","):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        if len(val_str.split(",")[-1]) == 2:
            val_str = val_str.replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def ejecutar(supabase_client, extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro):
    st.markdown("""
    <style>
    .titulo-principal { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black', sans-serif; }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 3px solid #0d1b2a !important; border-radius: 8px !important; overflow: hidden !important; }
    .hud-tarifas { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }
    .hud-tarifas-item { text-align: center; flex: 1; }
    .hud-tarifas-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-tarifas-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stSelectbox"] [data-baseweb="select"] { border: 2px solid #0d1b2a !important; border-radius: 6px !important; font-weight: 800 !important; }
    div[data-testid="stCodeBlock"] pre { border: 3px solid #0d1b2a !important; border-radius: 8px !important; }
    div[data-testid="stCodeBlock"] code { color: #0d1b2a !important; font-weight: 900 !important; font-size: 17px !important; font-family: 'Arial Black', monospace !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    
    if supabase_client is None:
        st.error("🚨 El enlace principal con Supabase no está inicializado.")
        return

    gc = inicializar_cliente_gspread()

    # --- 🧮 SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Computo de Perfiles)")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Descargando base central de Supabase y aplicando márgenes..."):
                try:
                    respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                    lista_precios = []
                    
                    if respuesta.data:
                        for row in respuesta.data:
                            prod = str(row.get("PRODUCTO", "")).strip().upper()
                            val_costo = row.get("COSTO", "0")
                            
                            if prod and prod != "PRODUCTO" and "INVENTARIO" not in prod:
                                costo_base = purificar_y_convertir_precio(val_costo)
                                
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
                        st.success(f"✅ ¡TARIFARIO DESPLEGADO! {len(lista_precios)} productos mapeados con sus márgenes.")
                    else:
                        st.error("🚨 Error matemático: Supabase respondió, pero la columna COSTO contiene formatos ilegibles o la tabla sigue vacía.")
                except Exception as e:
                    st.error(f"🚨 Falla crítica en descarga del Módulo 5: {e}")
                    
        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_t = st.session_state['df_tarifario']
            
            total_quimicos_tarifados = len(df_t)
            costo_maximo_comercial = df_t['TERCERO (+45.1%)'].max()
            costo_medio_base = df_t['COSTO BASE'].mean()
            
            st.markdown(f"""
            <div class="hud-tarifas">
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Insumos en Línea</p>
                    <p class="hud-tarifas-value">🧪 {total_quimicos_tarifados} ítems</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Costo Promedio</p>
                    <p class="hud-tarifas-value">💵 $ {costo_medio_base:,.0f}</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Tope Máximo</p>
                    <p class="hud-tarifas-value">📈 $ {costo_maximo_comercial:,.0f}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Consulta Multi-Producto"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes Oficiales")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### Caja de Copiado Masivo para SAP")
                col_margen = st.selectbox("Seleccione el Perfil de Productor:", 
                                          ["TERCERO (+45.1%)", "AFILIADO (+16.4%)", "COOPERATIVA / SOCIO (+11.2%)", "ORGÁNICO (+1.1%)", "COSTO BASE"])
                lista_textos = [fmt_sap(x) for x in df_t[col_margen]]
                st.code("\n".join(lista_textos), language="text")
                    
            with t3:
                st.markdown("#### Búsqueda Rápida de Costos y Márgenes")
                
                # 💥 LA EVOLUCIÓN: Mutación de selectbox a multiselect con precarga inteligente
                opciones_productos = df_t["PRODUCTO"].tolist()
                prods_sel = st.multiselect(
                    "🔍 Seleccione uno o varios Productos para comparar:", 
                    options=opciones_productos,
                    default=[opciones_productos[0]] if opciones_productos else []
                )
                
                # Despliegue en cascada controlada por cada elemento seleccionado
                for prod_sel in prods_sel:
                    datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
                    
                    st.markdown(f"#### 🧪 Arsenal: `{prod_sel}`")
                    c1, c2, c3, c4, c5 = st.columns(5)
                    caja_titulo = "height: 45px; display: flex; align-items: flex-end; margin-bottom: 5px;"
                    estilo_etiqueta = "font-size: 11px; font-weight: 900; color: #0d1b2a; margin: 0; line-height: 1.2;"
                    
                    with c1: 
                        st.markdown(f"<div style='{caja_titulo}'><p style='{estilo_etiqueta}'>🏷️ COSTO BASE</p></div>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["COSTO BASE"]))
                    with c2: 
                        st.markdown(f"<div style='{caja_titulo}'><p style='{estilo_etiqueta}'>🌱 ORGÁNICO<br>(+1.1%)</p></div>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["ORGÁNICO (+1.1%)"]))
                    with c3: 
                        st.markdown(f"<div style='{caja_titulo}'><p style='{estilo_etiqueta}'>🤝 SOCIO/COOP<br>(+11.2%)</p></div>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["COOPERATIVA / SOCIO (+11.2%)"]))
                    with c4: 
                        st.markdown(f"<div style='{caja_titulo}'><p style='{estilo_etiqueta}'>🏢 AFILIADO<br>(+16.4%)</p></div>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["AFILIADO (+16.4%)"]))
                    with c5: 
                        st.markdown(f"<div style='{caja_titulo}'><p style='{estilo_etiqueta}'>👤 TERCERO<br>(+45.1%)</p></div>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["TERCERO (+45.1%)"]))
                    st.markdown("<hr style='border:1px dashed #d4af37; margin-top:5px; margin-bottom:20px;'/>", unsafe_allow_html=True)

    # --- 🚀 SECCIÓN INFERIOR COMPLETA: OMEGA V12 ---
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    
    with st.container(border=True):
        c_url1, c_url2 = st.columns(2)
        with c_url1:
            st.text_input("🔗 1. Base de Origen Activa:", value="DATABASE: Supabase Cloud [PRECIOS_INSUMOS]", disabled=True)
        with c_url2:
            url_dest = st.text_input("🎯 2. URL de Sábana Destino (Google Sheets):", placeholder="Pegue el enlace completo aquí...")
        
        semana_target = st.number_input("🔢 Digite la Semana a actualizar (1 a 53):", min_value=1, max_value=53, value=24, step=1)
        
        c_btn1, c_btn2 = st.columns(2)
        
        with c_btn1:
            if st.button("📊 PREVISUALIZAR COMPORTAMIENTO DE PRECIOS POR DOSIS", use_container_width=True, type="secondary"):
                if gc is None:
                    st.error("🚨 Enlace satelital roto con Google Cloud.")
                elif not url_dest or "http" not in url_dest:
                    st.error("❌ Ingrese la URL de destino para previsualizar.")
                else:
                    try:
                        with st.spinner("🕵️‍♂️ Calculando Comportamiento Operativo..."):
                            respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                            dict_precios = {}
                            for row in respuesta.data:
                                prod = limpiar_texto_vba(row.get('PRODUCTO', '')).upper().strip()
                                precio_final = purificar_y_convertir_precio(row.get('COSTO', 0))
                                if prod and precio_final > 0:
                                    dict_precios[prod] = precio_final

                            sh_dest = gc.open_by_url(url_dest)
                            ws_datos = sh_dest.worksheet("DATOS")
                            datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                            
                            idx_fila_semanas = 6
                            for idx, r in enumerate(datos_dest[:12]):
                                r_str = [str(cell).strip().split('.')[0] for cell in r]
                                if any(w in r_str for w in ["11", "12", "13", "18"]):
                                    idx_fila_semanas = idx
                                    break
                            
                            filas_comp = []
                            for r_idx, row in enumerate(datos_dest):
                                n_fila = r_idx + 1
                                if n_fila < (idx_fila_semanas + 2): continue
                                
                                row_padded = row + [""] * (15 - len(row)) if len(row) < 15 else row
                                tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                                producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                                
                                if producto_dest in dict_precios:
                                    precio_pleno = dict_precios[producto_dest]
                                    dosis_valor = extraer_numero(row_padded[0])
                                    
                                    if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                        valor_dosis = precio_pleno * dosis_valor if dosis_valor > 0 else 0
                                        formula = f"{dosis_valor} Dosis × $ {precio_pleno:,.0f}"
                                    else:
                                        valor_dosis = precio_pleno
                                        formula = "Precio Unitario Directo"
                                        
                                    filas_comp.append({
                                        "Fila SAP": n_fila,
                                        "Tipo de Registro": tipo_tabla,
                                        "Insumo / Producto": producto_dest,
                                        "Precio Final Calculado": valor_dosis,
                                        "Lógica de Impacto": formula
                                    })
                            
                            if filas_comp:
                                df_vis = pd.DataFrame(filas_comp).copy()
                                df_vis["Precio Final Calculado"] = df_vis["Precio Final Calculado"].map("$ {:,.0f}".format).str.replace(",", ".")
                                st.dataframe(df_vis, use_container_width=True, hide_index=True)
                            else:
                                st.warning("⚠️ No se encontraron coincidencias.")
                    except Exception as e:
                        st.error(f"🚨 Falla en análisis: {e}")

            with c_btn2:
                if st.button("🚀 EJECUTAR SINCRONIZACIÓN OMEGA V12", use_container_width=True, type="primary"):
                    if gc is None:
                        st.error("🚨 Enlace satelital roto con Google Cloud.")
                        return
                    try:
                        with st.status("🕵️‍♂️ CONECTANDO CON CÉLULA SUPABASE...", expanded=True) as status:
                            respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                            dict_precios = {}
                            for row in respuesta.data:
                                prod = limpiar_texto_vba(row.get('PRODUCTO', '')).upper().strip()
                                precio_final = purificar_y_convertir_precio(row.get('COSTO', 0))
                                if prod and precio_final > 0:
                                    dict_precios[prod] = precio_final
                            
                            sh_dest = gc.open_by_url(url_dest)
                            ws_datos = sh_dest.worksheet("DATOS")
                            datos_dest = ws_datos.get_all_values(value_render_option='UNFORMATTED_VALUE')
                            
                            idx_fila_semanas = 6
                            for idx, r in enumerate(datos_dest[:12]):
                                r_str = [str(cell).strip().split('.')[0] for cell in r]
                                if any(w in r_str for w in ["11", "12", "13", "18"]):
                                    idx_fila_semanas = idx
                                    break
                            
                            col_semana = -1
                            for i, v in enumerate(datos_dest[idx_fila_semanas]):
                                if str(v).strip().split('.')[0] == str(semana_target):
                                    col_semana = i + 1
                                    break
                            
                            if col_semana == -1: col_semana = int(semana_target) + 5
                            
                            updates = [{'range': gspread.utils.rowcol_to_a1(idx_fila_semanas + 1, col_semana), 'values': [[int(semana_target)]]}]
                            
                            for r_idx, row in enumerate(datos_dest):
                                n_fila = r_idx + 1
                                if n_fila < (idx_fila_semanas + 2): continue
                                
                                row_padded = row + [""] * (max(col_semana + 2, 15) - len(row)) if len(row) < max(col_semana + 2, 15) else row
                                tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                                producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                                
                                if producto_dest in dict_precios:
                                    precio_unitario = dict_precios[producto_dest]
                                    if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                        dosis_valor = extraer_numero(row_padded[0])
                                        valor_final = precio_unitario * dosis_valor if dosis_valor > 0 else 0
                                    else:
                                        valor_final = precio_unitario
                                        
                                    updates.append({'range': gspread.utils.rowcol_to_a1(n_fila, col_semana), 'values': [[valor_final]]})

                            if len(updates) > 1:
                                ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                                status.update(label="🎯 ¡MÓDULO DE DOSIS AJUSTADO!", state="complete")
                                st.success(f"🎉 Precios impactados en la columna {col_semana} de la macro.")
                                st.balloons()
                    except Exception as e:
                        st.error(f"🚨 FALLA EN LA INYECCIÓN: {e}")

if __name__ == "__main__":
    pass
