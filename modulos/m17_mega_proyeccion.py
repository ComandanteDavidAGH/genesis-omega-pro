import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import gspread
import re
import math
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# =================================================================
# 🔌 MOTORES DE CONEXIÓN Y DESCARGA (Caché Optimizada)
# =================================================================

def obtener_cliente_gspread_unificado():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        return gspread.service_account(filename='credenciales.json')
    except: return None

@st.cache_data(show_spinner=False, ttl=600)
def cargar_boveda_mega_proyeccion():
    gc = obtener_cliente_gspread_unificado()
    if not gc: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {}
    
    df_mezclas, df_conf, df_dicc, df_t2, df_precios = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    hist_vuelo_promedio = {}

    try:
        # 1. Bóveda Recetas y Configuración
        boveda_recetas = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
        try:
            data_mez = boveda_recetas.worksheet("DD_Mesclas").get_all_values()
            df_mezclas = pd.DataFrame(data_mez[1:], columns=data_mez[0])
            df_mezclas['COCTEL_CLEAN'] = df_mezclas.iloc[:, 0].astype(str).str.upper().str.replace(" ", "")
        except: pass

        try: df_conf = pd.DataFrame(boveda_recetas.worksheet("Configuración").get_all_values()[1:], columns=boveda_recetas.worksheet("Configuración").get_all_values()[0])
        except: pass
        try: df_dicc = pd.DataFrame(boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()[1:], columns=boveda_recetas.worksheet("DICCIONARIO_SIGLAS").get_all_values()[0])
        except: pass
        try: 
            t2_raw = boveda_recetas.worksheet("TABLA 2").get_all_values()
            idx_t2 = next((i for i, r in enumerate(t2_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 0)
            df_t2 = pd.DataFrame(t2_raw[idx_t2+1:], columns=[str(c).strip() for c in t2_raw[idx_t2]])
        except: pass

        # 2. Bóveda Precios
        try:
            sh_precios = gc.open_by_url("https://docs.google.com/spreadsheets/d/1qZ4av-DH2oCJdgllBX27gdA2jEhT9bt2yv_sboORfSg/edit")
            precios_consolidados = []
            for ws in sh_precios.worksheets():
                datos_hoja = ws.get_all_values()
                if not datos_hoja: continue
                idx_header, col_anio, col_prod = -1, -1, -1
                for i in range(min(10, len(datos_hoja))):
                    fila_upper = [str(x).upper().strip() for x in datos_hoja[i]]
                    if 'AÑO' in fila_upper and 'PRODUCTO' in fila_upper:
                        idx_header, col_anio, col_prod = i, fila_upper.index('AÑO'), fila_upper.index('PRODUCTO'); break
                if idx_header != -1:
                    for row in datos_hoja[idx_header+1:]:
                        if len(row) > max(col_anio, col_prod):
                            anio_str, str_prod = str(row[col_anio]).strip().upper(), str(row[col_prod]).strip().upper()
                            if anio_str and str_prod:
                                vals = []
                                for v in row[max(col_anio, col_prod) + 1:]:
                                    val_c = re.sub(r'[^\d\.,\-]', '', str(v).strip())
                                    if val_c:
                                        if '.' in val_c and ',' in val_c: val_c = val_c.replace('.', '').replace(',', '.') if val_c.rfind(',') > val_c.rfind('.') else val_c.replace(',', '')
                                        elif ',' in val_c: val_c = val_c.replace(',', '.')
                                        try: vals.append(float(val_c))
                                        except: pass
                                if vals: precios_consolidados.append({'AÑO': anio_str, 'PRODUCTO': str_prod, 'PRODUCTO_CLEAN': str_prod.replace(" ", ""), 'PRECIO_PROM': sum(vals)/len(vals)})
            df_precios = pd.DataFrame(precios_consolidados)
        except: pass

        # 3. Inteligencia de Precio de Vuelo Histórico (Tabla 1) - ENFOCADO EN AÑO ACTUAL
        try:
            t1_raw = boveda_recetas.worksheet("TABLA 1").get_all_values()
            idx_t1 = next((i for i, r in enumerate(t1_raw) if "FINCA" in [str(x).upper().strip() for x in r]), 4)
            df_t1 = pd.DataFrame(t1_raw[idx_t1+1:], columns=t1_raw[idx_t1])
            
            col_finca = next((c for c in df_t1.columns if "FINCA" in str(c).upper()), None)
            col_ha = next((c for c in df_t1.columns if "AREA_FUMIG" in str(c).upper() or "CANTIDAD" in str(c).upper()), None)
            col_costo = next((c for c in df_t1.columns if "COSTO_HA" in str(c).upper() or "AVION" in str(c).upper()), None)
            col_fecha = next((c for c in df_t1.columns if "FECHA" in str(c).upper()), None)
            
            if col_finca and col_ha and col_costo and col_fecha:
                df_calc = df_t1[[col_finca, col_ha, col_costo, col_fecha]].copy()
                def limp_num(x):
                    try: return float(re.sub(r'[^\d\.]', '', str(x).replace(',', '.')))
                    except: return 0.0
                df_calc['HA'] = df_calc[col_ha].apply(limp_num)
                df_calc['COSTO'] = df_calc[col_costo].apply(limp_num)
                df_calc = df_calc[df_calc['HA'] > 0]
                
                año_actual = str(datetime.now().year) # Extrae 2026 automáticamente
                
                for finca, grp in df_calc.groupby(col_finca):
                    # Filtramos los vuelos de esta finca que ocurrieron en el año actual (2026)
                    grp_actual = grp[grp[col_fecha].astype(str).str.contains(año_actual)]
                    
                    if not grp_actual.empty:
                        # Si tiene vuelos este año, promediamos SOLO la realidad de este año
                        hist_vuelo_promedio[str(finca).strip().upper()] = grp_actual['COSTO'].mean()
                    else:
                        # Plan B: Si no ha volado este año, usamos su promedio histórico general
                        hist_vuelo_promedio[str(finca).strip().upper()] = grp['COSTO'].mean()
        except: pass
    except Exception as e: st.error(f"Error cargando bases: {e}")

    return df_mezclas, df_conf, df_dicc, df_precios, df_t2, hist_vuelo_promedio

# =================================================================
# 🧠 MOTORES DE LÓGICA (Márgenes y Volumetría)
# =================================================================

def limpiar_numero(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v = str(val).strip().replace(',', '.')
        v = re.sub(r'[^\d\.\-]', '', v)
        if v.count('.') > 1: v = v.rsplit('.', 1)[0].replace('.', '') + '.' + v.rsplit('.', 1)[1]
        return float(v) if v else 0.0
    except: return 0.0

def extraer_receta_mega(coctel_sel, finca_sel, df_mezclas, df_dicc, df_t2):
    coctel_u = str(coctel_sel).upper().strip().replace("+", " ").replace("-", " ")
    partes = coctel_u.split()
    base_coctel = partes[0] if partes else ""
    aditivos = partes[1:] if len(partes) > 1 else []
    
    dict_prods = {}
    es_organico = False
    try:
        if not df_t2.empty:
            match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_sel.upper().strip()]
            if not match_f.empty and "ORGANIC" in str(match_f.iloc[0, 5]).upper(): es_organico = True
    except: pass

    base_buscar = f"{base_coctel}O" if es_organico and not base_coctel.endswith('O') else base_coctel

    if not df_mezclas.empty:
        col_0 = df_mezclas.iloc[:, 0].astype(str).str.upper().str.strip()
        rb = df_mezclas[col_0 == base_buscar]
        if rb.empty and es_organico: rb = df_mezclas[col_0 == base_coctel]
        for _, r in rb.iterrows():
            p, d = str(r.iloc[1]).strip().upper(), limpiar_numero(r.iloc[2])
            if d > 0 and p not in ['NAN', 'NONE', '']: dict_prods[p] = d

    if not df_dicc.empty and aditivos:
        for ad in aditivos:
            m_s = df_dicc[df_dicc['SIGLA'].astype(str).str.upper().str.strip() == ad]
            if not m_s.empty:
                p_ad, d_ad = str(m_s.iloc[0]['PRODUCTO']).strip().upper(), limpiar_numero(m_s.iloc[0]['DOSIS'])
                if d_ad > 0 and p_ad not in ['NAN', 'NONE', '']: dict_prods[p_ad] = dict_prods.get(p_ad, 0.0) + d_ad

    for p in list(dict_prods.keys()):
        if "ACONDICIONADOR" in p: dict_prods[p] = 0.06 if any(x in coctel_u for x in ["ZN", "BT", "ZT", "ZITRON"]) else 0.02
        elif "IMBIOSIL" in p.replace(" ", ""): dict_prods[p] = 1.5 if base_coctel.startswith("IN") or "IMBIOSIL" in base_coctel else 1.0
        if es_organico and "ADHERENTE" in p: del dict_prods[p]
    if es_organico and not any("SPRAYFIX" in k for k in dict_prods.keys()): dict_prods["SPRAYFIX"] = 0.2
    
    return dict_prods

# =================================================================
# 👑 RENDERIZADO VISUAL
# =================================================================

def ejecutar():
    st.markdown("""
    <style>
    .titulo-mega { color: #0d1b2a; border-bottom: 3px solid #d4af37; padding-bottom: 5px; font-family: 'Arial Black'; margin-bottom: 15px;}
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] { border: 2px solid #0d1b2a !important; border-radius: 8px !important; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); overflow: hidden !important; }
    .tarjeta-kpi { background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%); border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white; box-shadow: 0px 4px 10px rgba(0,0,0,0.2); text-align: center; margin-bottom: 15px;}
    .kpi-titulo { font-size: 12px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .kpi-valor { font-size: 26px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-mega'>🚀 Módulo 17: Mega-Proyección Operativa</h1>", unsafe_allow_html=True)
    st.info("💡 **Guía Rápida:** Pega tus datos de Excel directamente en la tabla. Si dejas el `PRECIO VUELO` o las `HECTAREAS` en 0, el sistema buscará el dato oficial de esa finca automáticamente.")

    # 1. Cargar bases maestras en caché
    with st.spinner("Conectando con la Bóveda Maestra..."):
        df_mezclas, df_conf, df_dicc, df_precios, df_t2, hist_vuelo = cargar_boveda_mega_proyeccion()

    # Extracción de listas para los selectores
    lista_fincas = []
    if not df_t2.empty:
        lista_fincas = sorted([str(x).upper().strip() for x in df_t2.iloc[:, 0].dropna().unique() if str(x).upper().strip() not in ['NAN', 'NONE', '', 'FINCA', 'TOTAL']])
    
    lista_cocteles = []
    if not df_mezclas.empty:
        lista_cocteles = sorted([str(x).upper().strip() for x in df_mezclas.iloc[:, 0].dropna().unique() if str(x).upper().strip() not in ['NAN', 'NONE', '']])

    # 2. Configurar Data Editor Inicial
    if 'mega_input' not in st.session_state:
        st.session_state.mega_input = pd.DataFrame([{"FINCA": None, "HECTAREAS": 0.0, "COCTEL": None, "DIAS CICLO": 0, "PRECIO VUELO": 0.0} for _ in range(30)])

    # Escáner dinámico para Productor en TABLA 2
    col_prod_idx = 5
    if not df_t2.empty:
        for i, c_name in enumerate(df_t2.columns):
            if 'PROD' in str(c_name).upper() or 'TIPO' in str(c_name).upper(): col_prod_idx = i

    st.markdown("### 📥 1. Pista de Aterrizaje de Datos (Pegar desde Excel)")
    df_edited = st.data_editor(
        st.session_state.mega_input,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FINCA": st.column_config.SelectboxColumn("Finca", options=lista_fincas, required=True),
            "HECTAREAS": st.column_config.NumberColumn("Hectáreas", min_value=0.0, format="%.2f"),
            "COCTEL": st.column_config.SelectboxColumn("Cóctel", options=lista_cocteles),
            "DIAS CICLO": st.column_config.NumberColumn("Días Ciclo", min_value=0),
            "PRECIO VUELO": st.column_config.NumberColumn("Precio/Ha Vuelo", min_value=0.0, format="%.0f"),
        }
    )

    if st.button("🔥 EJECUTAR MEGA-PROYECCIÓN", type="primary", use_container_width=True):
        with st.spinner("Procesando matriz financiera y logística..."):
            resultados = []
            log_volumetrico = {}
            
            # Limpiar dataframe de filas vacías (evitar NoneType errors)
            df_valid = df_edited.dropna(subset=['FINCA']).copy()
            df_valid = df_valid[df_valid['FINCA'].astype(str).str.strip() != ""]
            
            año_actual = str(datetime.now().year)

            for idx, row in df_valid.iterrows():
                finca_n = str(row['FINCA']).strip().upper()
                ha_num = limpiar_numero(row['HECTAREAS'])
                coctel_n = str(row['COCTEL']).strip().upper() if pd.notna(row['COCTEL']) else ""
                dias_c = int(limpiar_numero(row['DIAS CICLO']))
                precio_vuelo = limpiar_numero(row['PRECIO VUELO'])

                # 💥 Auto-completar Hectáreas Oficiales si el usuario dejó 0
                if ha_num == 0 and not df_t2.empty:
                    match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                    if not match_f.empty:
                        ha_num = limpiar_numero(match_f.iloc[0].iloc[2]) # Columna de Área Bruta en TABLA 2

                if ha_num <= 0: continue

                # 💥 Auto-completar Precio Vuelo histórico si es 0
                if precio_vuelo == 0:
                    precio_vuelo = hist_vuelo.get(finca_n, 45000.0)

                # Inteligencia Productor y Márgenes
                tipo_prod = "TERCERO"
                if not df_t2.empty:
                    match_f = df_t2[df_t2.iloc[:, 0].astype(str).str.upper().str.strip() == finca_n]
                    if not match_f.empty: tipo_prod = str(match_f.iloc[0].iloc[col_prod_idx]).strip().upper() if len(match_f.columns) > col_prod_idx else "TERCERO"
                
                if "COOP" in finca_n or "EMPREBANCOOP" in finca_n: tipo_prod = "COOPERATIVA"

                mult_m, st_base, mult_v = 1.112, 1337.0, 1.112
                if not df_conf.empty:
                    match_cfg = df_conf[df_conf.iloc[:, 0].astype(str).str.strip().str.upper() == tipo_prod]
                    if not match_cfg.empty:
                        mult_m = limpiar_numero(match_cfg.iloc[0].iloc[3])
                        st_base = limpiar_numero(match_cfg.iloc[0].iloc[4])
                        mult_v = limpiar_numero(match_cfg.iloc[0].iloc[6])
                
                # Excepciones de backup si falla el excel
                if mult_m == 0:
                    if tipo_prod == "TERCERO": mult_m, st_base, mult_v = 1.451, 1583.0, 1.451
                    elif tipo_prod == "AFILIADO": mult_m, st_base, mult_v = 1.164, 1510.0, 1.164
                    elif tipo_prod == "COOPERATIVA": mult_m, st_base, mult_v = 1.112, 1510.0, 1.164
                    elif tipo_prod == "ORGANICO": mult_m, st_base, mult_v = 1.011, 1337.0, 1.011

                # Costo Mezcla
                costo_mezcla_fila = 0.0
                c_p_i, c_c_i = 8, 9
                if not df_conf.empty:
                    for i in range(min(5, len(df_conf))):
                        r_c = [str(x).upper() for x in df_conf.iloc[i]]
                        if 'PRODUCTO' in r_c and 'COSTO' in r_c:
                            c_p_i, c_c_i = r_c.index('PRODUCTO'), r_c.index('COSTO')
                            break

                dict_receta = extraer_receta_mega(coctel_n, finca_n, df_mezclas, df_dicc, df_t2)
                
                for p, d in dict_receta.items():
                    # Para la gráfica volumétrica
                    log_volumetrico[finca_n] = log_volumetrico.get(finca_n, {})
                    log_volumetrico[finca_n][p] = log_volumetrico[finca_n].get(p, 0.0) + (d * ha_num)

                    # Para Costos
                    precio_unitario = 0.0
                    if not df_conf.empty:
                        mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.strip() == p
                        if not mask_cfg.any() and "NEMATICIDA" in p: mask_cfg = df_conf.iloc[:, c_p_i].astype(str).str.upper().str.contains("NEMATI", na=False)
                        if mask_cfg.any(): precio_unitario = limpiar_numero(df_conf[mask_cfg].iloc[0, c_c_i])
                    
                    if precio_unitario == 0.0 and not df_precios.empty:
                        match_p = df_precios[(df_precios['AÑO'] == año_actual) & (df_precios['PRODUCTO_CLEAN'] == p.replace(" ",""))]
                        if not match_p.empty: precio_unitario = match_p['PRECIO_PROM'].mean()

                    costo_mezcla_fila += (d * ha_num * precio_unitario * mult_m)

                # Costos Totales
                costo_st_fila = dias_c * st_base * ha_num
                costo_vuelo_fila = precio_vuelo * ha_num # Directo y Neto
                gran_total = math.floor(costo_mezcla_fila + costo_st_fila + costo_vuelo_fila + 0.5)
                costo_ha = math.floor((gran_total / ha_num) + 0.5) if ha_num > 0 else 0

                resultados.append({
                    "FINCA": finca_n, "HECTAREAS": ha_num, "COCTEL": coctel_n, "DIAS CICLO": dias_c, "PRECIO VUELO": precio_vuelo,
                    "Costo ST ($)": math.floor(costo_st_fila), "Costo Vuelo ($)": math.floor(costo_vuelo_fila), "Costo Mezcla ($)": math.floor(costo_mezcla_fila),
                    "Costo x Ha ($)": costo_ha, "RESULTADO TOTAL ($)": gran_total
                })

            st.session_state.mega_resultados = pd.DataFrame(resultados)
            st.session_state.mega_volumetria = log_volumetrico
            st.success("✅ Proyección completada exitosamente.")

    # 3. RENDERIZADO DEL DASHBOARD (Si hay datos procesados)
    if 'mega_resultados' in st.session_state and not st.session_state.mega_resultados.empty:
        st.markdown("---")
        df_res = st.session_state.mega_resultados
        vol_dict = st.session_state.mega_volumetria
        
        fincas_procesadas = sorted(df_res['FINCA'].unique().tolist())
        
        st.markdown("### 🎛️ 2. Tablero de Mando y Filtros")
        fincas_filtro = st.multiselect("📍 Filtrar análisis por Finca(s) [Dejar vacío para ver TOTAL GENERAL]:", fincas_procesadas)
        
        # Filtrado Dinámico
        if fincas_filtro:
            df_filtro = df_res[df_res['FINCA'].isin(fincas_filtro)]
            vol_dict_filtro = {k: v for k, v in vol_dict.items() if k in fincas_filtro}
        else:
            df_filtro = df_res
            vol_dict_filtro = vol_dict

        # Consolidación Volumétrica Filtrada
        cons_vol_agrupado = {}
        for f, prods in vol_dict_filtro.items():
            for p, vol in prods.items():
                cons_vol_agrupado[p] = cons_vol_agrupado.get(p, 0.0) + vol
        
        # Tarjetas de KPI
        t_st = df_filtro['Costo ST ($)'].sum()
        t_vu = df_filtro['Costo Vuelo ($)'].sum()
        t_mx = df_filtro['Costo Mezcla ($)'].sum()
        t_gr = df_filtro['RESULTADO TOTAL ($)'].sum()

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>👨‍🔬 Total Serv. Tec</p><p class='kpi-valor'>$ {t_st:,.0f}</p></div>".replace(",", "."), unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>✈️ Total Vuelo</p><p class='kpi-valor'>$ {t_vu:,.0f}</p></div>".replace(",", "."), unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='tarjeta-kpi'><p class='kpi-titulo'>🧪 Total Mezcla</p><p class='kpi-valor'>$ {t_mx:,.0f}</p></div>".replace(",", "."), unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='tarjeta-kpi' style='border-left: 5px solid #00ff00;'><p class='kpi-titulo' style='color:#00ff00;'>🔥 GRAN TOTAL</p><p class='kpi-valor'>$ {t_gr:,.0f}</p></div>".replace(",", "."), unsafe_allow_html=True)

        # Tablas y Gráficas
        tab1, tab2 = st.tabs(["📊 Detalles Económicos Fila x Fila", "📦 Auditoría Volumétrica de Insumos"])
        
        with tab1:
            df_view = df_filtro.copy()
            for col in ["PRECIO VUELO", "Costo ST ($)", "Costo Vuelo ($)", "Costo Mezcla ($)", "Costo x Ha ($)", "RESULTADO TOTAL ($)"]:
                df_view[col] = df_view[col].map("$ {:,.0f}".format).str.replace(",", "X").str.replace(".", ",").str.replace("X", ".")
            st.dataframe(df_view, use_container_width=True, hide_index=True)

        with tab2:
            if cons_vol_agrupado:
                df_insumos = pd.DataFrame(list(cons_vol_agrupado.items()), columns=["🧪 PRODUCTO", "VOLUMEN ESTIMADO"]).sort_values("VOLUMEN ESTIMADO", ascending=False)
                df_insumos["📦 VOLUMEN ESTIMADO (L/Kg)"] = df_insumos["VOLUMEN ESTIMADO"].map("{:,.1f}".format).str.replace(",", "X").str.replace(".", ",").str.replace("X", ".")
                
                c_tbl, c_grf = st.columns([1, 1.2])
                with c_tbl:
                    st.dataframe(df_insumos[["🧪 PRODUCTO", "📦 VOLUMEN ESTIMADO (L/Kg)"]], use_container_width=True, hide_index=True)
                with c_grf:
                    df_grafica = df_insumos.head(15).copy()
                    fig = px.bar(
                        df_grafica, y="🧪 PRODUCTO", x="VOLUMEN ESTIMADO", text="📦 VOLUMEN ESTIMADO (L/Kg)",
                        orientation='h', color="VOLUMEN ESTIMADO", color_continuous_scale="GnBu",
                        title=f"Top 15 Insumos Proyectados"
                    )
                    fig.update_traces(textposition='outside', textfont_size=12)
                    fig.update_layout(yaxis={'categoryorder':'total ascending'}, plot_bgcolor='rgba(0,0,0,0)', margin=dict(r=100))
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No hay datos de insumos químicos para las fincas seleccionadas.")

        # ==========================================
        # 💾 SÚPER EXPORTACIÓN A EXCEL (3 Pestañas)
        # ==========================================
        st.markdown("<br>", unsafe_allow_html=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtro.to_excel(writer, sheet_name='Detalle_Económico', index=False)
            if cons_vol_agrupado:
                df_insumos[["🧪 PRODUCTO", "VOLUMEN ESTIMADO"]].to_excel(writer, sheet_name='Consumo_Insumos', index=False)
            
            workbook = writer.book
            borde = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                ws.sheet_view.showGridLines = False
                for col_idx in range(1, ws.max_column + 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 20
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = borde
                        if cell.row == 1:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(vertical='center')
                            col_name = str(ws.cell(row=1, column=cell.column).value).upper()
                            if isinstance(cell.value, (int, float)):
                                if "COSTO" in col_name or "PRECIO" in col_name or "RESULTADO" in col_name or "TOTAL" in col_name:
                                    cell.number_format = '"$" #,##0' 
                                elif "HECTAREAS" in col_name or "VOLUMEN" in col_name:
                                    cell.number_format = '#,##0.0'

        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL)",
            data=buffer.getvalue(),
            file_name=f"MegaProyeccion_Operativa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    pass
