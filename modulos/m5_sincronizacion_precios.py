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
    except:
        return None

# =================================================================
# 👑 PROCESAMIENTO PRINCIPAL DE TARIFAS Y MACRO OMEGA V12
# =================================================================

def ejecutar(extraer_numero, fmt_sap, limpiar_texto_vba, val_seguro):
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
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Sincronización de Precios y Tarifas</h1>", unsafe_allow_html=True)
    
    gc = inicializar_cliente_gspread()
    if gc is None:
        st.error("🚨 Enlace satelital roto con Google Cloud. Verifique sus credenciales.")
        return

    # --- 🧮 SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Copia Rápida)")
        st.info("💡 Obtenga la lista de precios exactos multiplicados por el margen de cada perfil, listos para copiar y pegar en SAP.")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Conectando con la Bóveda de Configuración a alta velocidad..."):
                try:
                    url_gen = "https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit"
                    sh_gen = gc.open_by_url(url_gen)
                    raw_config = sh_gen.worksheet("Configuración").get_all_values()
                    
                    lista_precios = []
                    for row in raw_config:
                        if len(row) > 9:
                            prod = str(row[8]).upper().strip()
                            
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
                        st.success(f"✅ Tarifario cargado en la caché local: {len(lista_precios)} productos.")
                    else:
                        st.warning("⚠️ El escáner no encontró productos con precios válidos en la hoja.")
                except Exception as e:
                    st.error(f"🚨 Error al generar tarifario: {e}")
                    
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
                    
                    # 🎨 Estilo forzado: Letra oscura, gruesa y fácil de leer
                    estilo_etiqueta = "font-size: 13px; font-weight: 900; color: #0d1b2a; margin-bottom: 2px; letter-spacing: 0.5px;"
                    
                    with c1: 
                        st.markdown(f"<p style='{estilo_etiqueta}'>🏷️ COSTO BASE</p>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["COSTO BASE"]))
                    with c2: 
                        st.markdown(f"<p style='{estilo_etiqueta}'>🌱 ORGÁNICO (+1.1%)</p>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["ORGÁNICO (+1.1%)"]))
                    with c3: 
                        st.markdown(f"<p style='{estilo_etiqueta}'>🤝 SOCIO / COOP (+11.2%)</p>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["COOPERATIVA / SOCIO (+11.2%)"]))
                    with c4: 
                        st.markdown(f"<p style='{estilo_etiqueta}'>🏢 AFILIADO (+16.4%)</p>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["AFILIADO (+16.4%)"]))
                    with c5: 
                        st.markdown(f"<p style='{estilo_etiqueta}'>👤 TERCERO (+45.1%)</p>", unsafe_allow_html=True)
                        st.code(fmt_sap(datos_prod["TERCERO (+45.1%)"]))
                        
    st.markdown("---")
    st.markdown("### 🚀 Sincronización Automática a la Macro (Omega V12)")
    
    c_url1, c_url2 = st.columns(2)
    with c_url1:
        url_ori = st.text_input("🔗 1. URL de Bóveda Origen (GÉNESIS_CONFIG):", value="https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
    with c_url2:
        url_dest = st.text_input("🎯 2. URL de Sábana Destino (Google Sheets Nativo convertido):", placeholder="Pegue aquí el enlace completo de 44 caracteres...")
    
    semana_target = st.number_input("🔢 Digite la Semana a actualizar (1 a 53):", min_value=1, max_value=53, value=24, step=1)

    if st.button("🚀 EJECUTAR OMEGA V12", use_container_width=True):
        if not url_ori or not url_dest or "http" not in url_dest:
            st.error("❌ Por favor, asegúrese de ingresar ambas URLs válidas en la pantalla antes de disparar.")
            return
            
        try:
            with st.status("🕵️‍♂️ ANALIZANDO CONEXIÓN PERIMETRAL EN VIVO...", expanded=True) as status:
                
                sh_gen = gc.open_by_url(url_ori)
                raw_config = sh_gen.worksheet("Configuración").get_all_values(value_render_option='UNFORMATTED_VALUE')
                
                dict_precios = {}
                for row in raw_config:
                    if len(row) > 9:
                        prod = limpiar_texto_vba(row[8]).upper().strip()
                        if prod and prod != "PRODUCTO":
                            dict_precios[prod] = val_seguro(row[9])
                
                st.write(f"📊 **Origen:** `{len(dict_precios)}` precios leídos de Configuración.")

                # Conexión directa a la sábana destino
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
                    if n_fila < (idx_fila_semanas + 2): continue
                    
                    row_padded = row + [""] * (max(col_semana + 2, 15) - len(row)) if len(row) < max(col_semana + 2, 15) else row
                    tipo_tabla = limpiar_texto_vba(row_padded[1]).upper().strip() 
                    producto_dest = limpiar_texto_vba(row_padded[3]).upper().strip()
                    
                    if not producto_dest: continue
                    
                    if producto_dest in dict_precios:
                        precio_unitario = dict_precios[producto_dest]
                        
                        # 🎯 LA JUGADA MAESTRA: Si pertenece a la tabla de dosis por hectárea
                        if "DOSIS-HA" in tipo_tabla.replace(" ", ""):
                            # Extraemos el factor numérico directamente de la Columna A (índice 0) de la misma fila
                            dosis_valor = extraer_numero(row_padded[0])
                            if dosis_valor > 0:
                                valor_final = precio_unitario * dosis_valor
                            else:
                                valor_final = 0 # Imita fielmente tu =SI.ERROR(..., 0) si la celda de dosis está vacía o es texto
                        else:
                            # Primera tabla: Precio por Litro puro
                            valor_final = precio_unitario
                            
                        updates.append({
                            'range': gspread.utils.rowcol_to_a1(n_fila, col_semana),
                            'values': [[valor_final]]
                        })

                if len(updates) > 1:
                    ws_datos.batch_update(updates, value_input_option='USER_ENTERED')
                    status.update(label="🎯 ¡MÓDULO DE DOSIS AJUSTADO AL 100%!", state="complete")
                    st.success(f"🎉 Los precios unitarios y por dosis calculados desde la Columna A han impactado en la columna {col_semana}.")
                    st.balloons()
                else:
                    status.update(label="❌ OPERACIÓN SIN COINCIDENCIAS", state="error")
                    st.error("No se generaron actualizaciones.")

        except Exception as e:
            st.error(f"🚨 FALLA TRANSACCIONAL: {e}")

if __name__ == "__main__":
    pass
