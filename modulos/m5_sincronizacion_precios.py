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
    """ Convierte cualquier formato de texto de SAP/Excel a un número flotante real """
    if not valor_crudo:
        return 0.0
    
    # Limpieza de caracteres cosméticos
    val_str = str(valor_crudo).replace("$", "").replace("COP", "").replace(" ", "").strip()
    
    # Manejo de formatos de miles y decimales cruzados (Ej: 11.953,50 o 11,953.50)
    if "," in val_str and "." in val_str:
        if val_str.find(".") < val_str.find(","):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        # Si tiene solo comas, asumimos formato decimal o miles según el contexto de SAP
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

    # --- 🧮 SECCIÓN: TARIFARIO MAESTRO ---
    with st.container(border=True):
        st.markdown("### 🧮 Tarifario Maestro Dinámico (Visor y Computo de Perfiles)")
        
        if st.button("🔄 Cargar / Actualizar Tarifario Maestro", type="secondary", use_container_width=True):
            with st.spinner("📡 Descargando base central y aplicando márgenes de perfiles..."):
                try:
                    respuesta = supabase_client.table("PRECIOS_INSUMOS").select("*").execute()
                    lista_precios = []
                    
                    if respuesta.data:
                        for row in respuesta.data:
                            prod = str(row.get("PRODUCTO", "")).strip().upper()
                            val_costo = row.get("COSTO", "0")
                            
                            if prod and prod != "PRODUCTO" and "INVENTARIO" not in prod:
                                # Aplicamos el purificador numérico avanzado para romper el estancamiento
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
                        st.success(f"✅ ¡TARIFARIO DESPLEGADO! {len(lista_precios)} productos mapeados con sus respectivos márgenes comerciales.")
                    else:
                        st.error("🚨 Error matemático: Supabase respondió, pero la columna COSTO contiene formatos ilegibles o la tabla está vacía.")
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
                    <p class="hud-tarifas-title">Costo Promedio Mercado</p>
                    <p class="hud-tarifas-value">💵 $ {costo_medio_base:,.0f}</p>
                </div>
                <div class="hud-tarifas-item">
                    <p class="hud-tarifas-title">Tope Máximo de Venta</p>
                    <p class="hud-tarifas-value">📈 $ {costo_maximo_comercial:,.0f}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            t1, t2, t3 = st.tabs(["💰 Visor General del Arsenal", "📋 Copia Masiva (Por Margen)", "🎯 Copia Individual"])
            
            with t1:
                st.markdown("#### Matriz de Costos y Márgenes Oficiales")
                df_visual = df_t.copy()
                for col in df_visual.columns:
                    if col != "PRODUCTO":
                        df_visual[col] = df_visual[col].map("$ {:,.0f}".format).str.replace(",", ".")
                st.dataframe(df_visual, use_container_width=True, hide_index=True)
                
            with t2:
                st.markdown("#### Caja de Copiado Masivo para SAP")
                col_margen = st.selectbox("Seleccione el Perfil de Productor para exportar:", 
                                          ["TERCERO (+45.1%)", "AFILIADO (+16.4%)", "COOPERATIVA / SOCIO (+11.2%)", "ORGÁNICO (+1.1%)", "COSTO BASE"])
                
                lista_textos = [fmt_sap(x) for x in df_t[col_margen]]
                st.code("\n".join(lista_textos), language="text")
                    
            with t3:
                st.markdown("#### Búsqueda Rápida Individual")
                prod_sel = st.selectbox("🔍 Seleccione Producto:", df_t["PRODUCTO"].tolist())
                if prod_sel:
                    datos_prod = df_t[df_t["PRODUCTO"] == prod_sel].iloc[0]
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

if __name__ == "__main__":
    pass
