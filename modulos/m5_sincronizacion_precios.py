import streamlit as st
import pandas as pd
import gspread
import re
import math
import io

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================
@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    """ Centraliza la autenticación con Google Cloud una sola vez en RAM """
    try:
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gspread.service_account(filename='credenciales.json')
    except Exception:
        return None

# =================================================================
# 💥 RASTREADOR OMNIDIRECCIONAL DE PRECIOS Y PROTECTOR DE DECIMALES
# =================================================================
def rastrear_precio_real(row):
    # 1. Búsqueda agresiva: Prioridad a columnas PRECIO, ACTUAL o SAP
    for k, v in row.items():
        k_up = str(k).upper()
        if "PRECIO" in k_up or "SAP" in k_up or "ACTUAL" in k_up:
            v_str = str(v).replace("$", "").replace("COP", "").strip()
            if "," in v_str and "." in v_str: 
                v_str = v_str.replace(".", "").replace(",", ".")
            elif "," in v_str: 
                v_str = v_str.replace(",", ".")
            try: 
                num = float(v_str)
                if num > 0: return num
            except: pass
    
    # 2. Búsqueda de respaldo: Si no hay precios nuevos, busca COSTO
    for k, v in row.items():
        if "COSTO" in str(k).upper():
            v_str = str(v).replace("$", "").replace("COP", "").strip()
            if "," in v_str and "." in v_str: 
                v_str = v_str.replace(".", "").replace(",", ".")
            elif "," in v_str: 
                v_str = v_str.replace(",", ".")
            try: 
                num = float(v_str)
                if num > 0: return num
            except: pass
    return 0.0

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE TARIFAS Y MACRO OMEGA V12
# =================================================================
def ejecutar(supabase_client, extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro):
    # 🚀 REFORZAMIENTO ESTÉTICO VIP COMPLETO
    st.markdown("""
    <style>
    .titulo-principal { 
        color: #0d1b2a; 
        border-bottom: 3px solid #d4af37; 
        padding-bottom: 5px; 
        font-family: 'Arial Black', sans-serif; 
    }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { 
        border: 3px solid #0d1b2a !important; 
        border-radius: 8px !important; 
        overflow: hidden !important; 
    }
    .hud-tarifas {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .hud-tarifas-item { text-align: center; flex: 1; }
    .hud-recargos-title, .hud-tarifas-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-tarifas-value { font-size: 22px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    
    /* 💥 DETONACIÓN DE CONTROLES PÁLIDOS */
    div[data-testid="stTextInput"] input, 
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] [data-baseweb="select"] {
        border: 2px solid #0d1b2a !important;
        border-radius: 6px !important;
        background-color: #ffffff !important;
        color: #0d1b2a !important;
        font-weight: 800 !important;
        font-size: 15px !important;
    }
    
    div[data-testid="stCodeBlock"], 
    div[data-testid="stCodeBlock"] pre, 
    div[data-testid="stCodeBlock"] pre code {
        background-color: #ffffff !important;
        border: 3px solid #0d1b2a !important;
        border-radius: 8px !important;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.08) !important;
        overflow: hidden !important;
        padding: 2px 5px !important;
    }
    div[data-testid="stCodeBlock"] code,
    div[data-testid="stCodeBlock"] code span,
    div[data-testid="stCodeBlock"] pre span {
        color: #0d1b2a !important;
        font-weight: 900 !important;
        font-size: 17px !important;
        font-family: 'Arial Black', monospace !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    
    if supabase_client is None:
        st.error("🚨 El enlace principal con la base de datos Supabase no está inicializado.")
        return

    gc = inicializar_cliente_gspread()

    # --- 🧮 SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Copia Rápida)")
        st.info("💡 Obtenga la lista de precios exactos multiplicados por el margen de cada perfil, listos para copiar y pegar en SAP.")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Descargando arsenal de precios desde Supabase Cloud..."):
                try:
                    respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                    lista_precios = []
                    
                    for row in respuesta.data:
                        prod = str(row.get('PRODUCTO', row.get('producto', ''))).upper().strip()
                        
                        es_cero_basura = False
                        try:
                            if float(prod) == 0: es_cero_basura = True
                        except ValueError:
                            pass
                            
                        if prod and prod != "PRODUCTO" and "INVENTARIO" not in prod and not es_cero_basura:
                            # 💥 LLAMADO AL RASTREADOR OMNIDIRECCIONAL
                            costo_base = rastrear_precio_real(row)
                            
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
                        st.success(f"✅ Tarifario cargado con éxito: {len(lista_precios)} productos extraídos.")
                    else:
                        st.warning("⚠️ No se pudieron procesar los valores numéricos correctamente.")
                except Exception as e:
                    st.error(f"🚨 Error al consultar Supabase: {e}")
                    
        if 'df_tarifario' in st.session_state and not st.session_state['df_tarifario'].empty:
            df_t = st.session_state['df_tarifario']
            
            total_quimicos_tarifados = len(df_t)
            costo_maximo_comercial = df_t['TERCERO (+45.1%)'].max()
            costo_medio_base = df_t['COSTO BASE'].mean()
            
            st.markdown(f"""
            <div class="hud-tarifas">
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Insumos Activos en Matriz</p>
                    <p class="hud-tarifas-value">🧪 {total_quimicos_tarifados} Productos</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Costo Promedio Base</p>
                    <p class="hud-tarifas-value">💵 $ {costo_medio_base:,.0f}</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Tope Máximo Tercero</p>
                    <p class="hud-tarifas-value">📈 $ {costo_maximo_comercial:,.0f}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Copia Individual (Por Producto)"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes (Ordenada por Producto)")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### Caja de Copiado Masivo")
                col_margen = st.selectbox("1️⃣ Seleccione el Perfil de Productor:", 
                                          ["TERCERO (+45.1%)", "AFILIADO (+16.4%)", "COOPERATIVA / SOCIO (+11.2%)", "ORGÁNICO (+1.1%)", "COSTO BASE"])
                incluir_nombres = st.toggle("🔘 Incluir Nombre del Producto (Alineación Perfecta)", value=False)
                
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
                st.markdown("#### Búsqueda Rápida Individual")
                prod_sel = st.selectbox("🔍 Buscar Producto Específico:", df_t["PRODUCTO"].tolist())
                if prod_sel:
                    datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
                    st.info(f"🎯 Valores calculados para: **{prod_sel}**")
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
                        
    # --- 🚀 SECCIÓN: SINCRONIZACIÓN Y MATRIZ DE COMPORTAMIENTO POR DOSIS ---
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    
    with st.container(border=True):
        c_url1, c_url2 = st.columns(2)
        with c_url1:
            st.text_input("🔗 1. Base de Origen Activa:", value="DATABASE: Supabase Cloud [PRECIOS_INSUMOS]", disabled=True)
        with c_url2:
            url_dest = st.text_input("🎯 2. URL de Sábana Destino (Google Sheets Nativo convertido):", placeholder="Pegue aquí el enlace completo de 44 caracteres...")
        
        semana_target = st.number_input("🔢 Digite la Semana a actualizar (1 a 53):", min_value=1, max_value=53, value=24, step=1)
        
        c_btn1, c_btn2 = st.columns(2)
        
        with c_btn1:
            if st.button("📊 PREVISUALIZAR COMPORTAMIENTO DE PRECIOS POR DOSIS", use_container_width=True, type="secondary"):
                if gc is None:
                    st.error("🚨 Enlace satelital roto con Google Cloud (gspread).")
                elif not url_dest or "http" not in url_dest:
                    st.error("❌ Ingrese la URL de la Sábana Destino para previsualizar los cálculos.")
                else:
                    try:
                        with st.spinner("🕵️‍♂️ Calculando Comportamiento Operativo..."):
                            respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                            dict_precios = {}
                            for row in respuesta.data:
                                prod = limpiar_texto_vba(row.get('PRODUCTO', row.get('producto', ''))).upper().strip()
                                # 💥 LLAMADO AL RASTREADOR OMNIDIRECCIONAL
                                precio_final = rastrear_precio_real(row)
                                
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
                                if n_fila < (idx_fila_semanas + 2):
                                    continue
                                
                                row_padded = row + [""] * (15 - len(row)) if len(row) < 15 else row
                                tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                                producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                                
                                if not producto_dest:
                                    continue
                                
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
                                        "Dosis (Col A)": dosis_valor if "DOSIS-HA" in tipo_tabla.replace(" ", "") else 1.0,
                                        "Precio Pleno (Origen)": precio_pleno,
                                        "Precio Final Calculado": valor_dosis,
                                        "Lógica de Impacto": formula
                                    })
                            
                            if filas_comp:
                                df_comp = pd.DataFrame(filas_comp)
                                st.markdown(f"#### 📉 Comportamiento Táctico de Precios — Semana {semana_target}")
                                df_vis = df_comp.copy()
                                df_vis["Precio Pleno (Origen)"] = df_vis["Precio Pleno (Origen)"].map("$ {:,.0f}".format).str.replace(",", ".")
                                df_vis["Precio Final Calculado"] = df_vis["Precio Final Calculado"].map("$ {:,.0f}".format).str.replace(",", ".")
                                st.dataframe(df_vis, use_container_width=True, hide_index=True)
                                st.success(f"📋 Análisis completo: {len(df_comp)} registros listos para inyección.")
                            else:
                                st.warning("⚠️ No se encontraron coincidencias entre Supabase y la Sábana.")
                    except Exception as e:
                        st.error(f"🚨 Falla en el análisis de comportamiento: {e}")

        with c_btn2:
            if st.button("🚀 EJECUTAR SINCRONIZACIÓN OMEGA V12", use_container_width=True, type="primary"):
                if gc is None:
                    st.error("🚨 Enlace satelital roto con Google Cloud (gspread).")
                    return
                if not url_dest or "http" not in url_dest:
                    st.error("❌ Por favor, ingrese una URL de destino válida para inyectar los datos.")
                    return
                    
                try:
                    with st.status("🕵️‍♂️ CONECTANDO CON CÉLULA SUPABASE Y DESTINO...", expanded=True) as status:
                        respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                        dict_precios = {}
                        for row in respuesta.data:
                            prod = limpiar_texto_vba(row.get('PRODUCTO', row.get('producto', ''))).upper().strip()
                            # 💥 LLAMADO AL RASTREADOR OMNIDIRECCIONAL
                            precio_final = rastrear_precio_real(row)
                            
                            if prod and precio_final > 0:
                                dict_precios[prod] = precio_final
                        
                        st.write(f"📊 **Supabase:** `{len(dict_precios)}` precios maestros mapeados.")

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
                            v_limpio = str(v).strip().split('.')[0]
                            if v_limpio == str(semana_target):
                                col_semana = i + 1
                                break
                        
                        if col_semana == -1:
                            col_semana = int(semana_target) + 5
                        
                        updates = []
                        updates.append({
                            'range': gspread.utils.rowcol_to_a1(idx_fila_semanas + 1, col_semana),
                            'values': [[int(semana_target)]]
                        })
                        
                        for r_idx, row in enumerate(datos_dest):
                            n_fila = r_idx + 1
                            if n_fila < (idx_fila_semanas + 2):
                                continue
                            
                            row_padded = row + [""] * (max(col_semana + 2, 15) - len(row)) if len(row) < max(col_semana + 2, 15) else row
                            tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                            producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                            
                            if not producto_dest:
                                continue
                            
                            if producto_dest in dict_precios:
                                precio_unitario = dict_precios[producto_dest]
                                
                                if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                                    dosis_valor = extraer_numero(row_padded[0])
                                    valor_final = precio_unitario * dosis_valor if dosis_valor > 0 else 0
                                else:
                                    valor_final = precio_unitario
                                    
                                updates.append({
                                    'range': gspread.utils.rowcol_to_a1(n_fila, col_semana),
                                    'values': [[valor_final]]
                                })

                        if len(updates) > 1:
                            ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                            status.update(label="🎯 ¡MÓDULO DE DOSIS AJUSTADO DESDE SUPABASE!", state="complete")
                            st.success(f"🎉 Precios impactados con éxito en la columna {col_semana}.")
                            st.balloons()
                        else:
                            status.update(label="❌ OPERACIÓN SIN COINCIDENCIAS", state="error")
                            st.error("No se generaron actualizaciones.")

                except Exception as e:
                    st.error(f"🚨 FALLA EN LA INYECCIÓN: {e}")

if __name__ == "__main__":
    pass
