import streamlit as st
import pandas as pd
import io

def ejecutar(quitar_tildes, purificar_lote):
    st.markdown("<h1 class='titulo-principal'>⚖️ Arqueo de Inventarios y Conciliación</h1>", unsafe_allow_html=True)
    
    # 📦 ZONA DE CARGA EN PANTALLA PRINCIPAL (Descamuflada)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### 📁 1. Sábana SAP")
        archivo_sap = st.file_uploader("1️⃣ Sábana de SAP", type=['xlsx', 'csv'])
    with c2:
        st.markdown("### 📋 2. Reportes Físicos")
        archivos_sup = st.file_uploader("2️⃣ Reportes Supervisores (.xlsx)", type=['xlsx'], accept_multiple_files=True)
    with c3:
        st.markdown("### 🎯 3. Objetivo")
        semana_obj = st.text_input("Semana a Auditar (Ej: 17):", placeholder="Escriba la semana aquí...")

    if "arqueo_procesado" not in st.session_state:
        st.session_state.arqueo_procesado = False
    if "observaciones_memoria" not in st.session_state:
        st.session_state.observaciones_memoria = {}

    def generar_cruce():
        cruce = pd.merge(st.session_state.df_sap_grouped, st.session_state.df_sup_grouped, on=['PISTA', 'LOTE_KEY'], how='outer')
        
        cruce['ITEM'] = cruce['ITEM'].fillna("---")
        cruce['PRODUCTO'] = cruce['PRODUCTO'].fillna(cruce['PRODUCTO_SUP']).fillna("N/A")
        cruce['LOTE'] = cruce['LOTE'].fillna(cruce['LOTE_SUP'])
        cruce['SALDO_SAP'] = cruce['SALDO_SAP'].fillna(0).round(2)
        cruce['SALDO_FISICO'] = cruce['SALDO_FISICO'].fillna(0).round(2)
        
        cruce = cruce[~((cruce['SALDO_SAP'] == 0) & (cruce['SALDO_FISICO'] == 0))]
        
        cruce['DIFERENCIA'] = (cruce['SALDO_FISICO'] - cruce['SALDO_SAP']).round(2)
        cruce['ESTADO'] = cruce['DIFERENCIA'].apply(lambda x: "✅ OK" if abs(x) <= 0.05 else "❌ DISCREPANCIA")
        
        cruce['OBSERVACIONES'] = ""
        for idx, row in cruce.iterrows():
            key = f"{row['PISTA']}_{row['LOTE_KEY']}"
            if key in st.session_state.observaciones_memoria:
                cruce.at[idx, 'OBSERVACIONES'] = st.session_state.observaciones_memoria[key]
            elif row['SALDO_SAP'] > 0 and row['SALDO_FISICO'] == 0:
                cruce.at[idx, 'OBSERVACIONES'] = "SUGERIDO: Entrega / Traslado / Pendiente por Facturar"

        st.session_state.cruce_final = cruce[['PISTA', 'ITEM', 'PRODUCTO', 'LOTE_KEY', 'LOTE', 'SALDO_SAP', 'SALDO_FISICO', 'DIFERENCIA', 'ESTADO', 'OBSERVACIONES']].sort_values(by=['PISTA', 'PRODUCTO'])

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Botón en la pantalla principal
    if st.button("🚀 INICIAR ARQUEO ESTRATÉGICO", type="primary", use_container_width=True):
        if not archivo_sap or not archivos_sup or not semana_obj:
            st.error("❌ Faltan suministros. Asegúrese de cargar ambos archivos y escribir la semana.")
        else:
            try:
                with st.spinner("Desplegando analista de inventarios..."):
                    st.session_state.observaciones_memoria = {}
                    
                    sap_file = archivo_sap[0] if isinstance(archivo_sap, list) else archivo_sap
                    nombre_sap = sap_file.name.lower()
                    if nombre_sap.endswith('.xlsx') or nombre_sap.endswith('.xls'):
                        df_sap = pd.read_excel(sap_file)
                    else:
                        try:
                            df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='utf-8')
                        except UnicodeDecodeError:
                            sap_file.seek(0)
                            df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='latin1')

                    df_sap.columns = [quitar_tildes(c) for c in df_sap.columns]
                    
                    c_item = next((c for c in df_sap.columns if "MATERIAL" in c and "DESC" not in c), df_sap.columns[0])
                    c_desc = next((c for c in df_sap.columns if "DESCRIP" in c), df_sap.columns[1])
                    c_pista = next((c for c in df_sap.columns if "ALMACEN" in c or "PISTA" in c), df_sap.columns[2])
                    c_lote = next((c for c in df_sap.columns if "LOTE" in c), df_sap.columns[3])
                    c_saldo = next((c for c in df_sap.columns if "LIBRE" in c or "UTILIZACION" in c), df_sap.columns[4])

                    df_sap_clean = df_sap[[c_item, c_desc, c_pista, c_lote, c_saldo]].copy()
                    df_sap_clean.columns = ['ITEM', 'PRODUCTO', 'PISTA', 'LOTE', 'SALDO_SAP']
                    df_sap_clean['LOTE_KEY'] = df_sap_clean['LOTE'].apply(purificar_lote)
                    df_sap_clean['PISTA'] = df_sap_clean['PISTA'].astype(str).str.strip().str.upper()
                    df_sap_clean['SALDO_SAP'] = pd.to_numeric(df_sap_clean['SALDO_SAP'].astype(str).replace(',', '.'), errors='coerce').fillna(0)
                    
                    st.session_state.df_sap_raw = df_sap_clean 
                    st.session_state.df_sap_grouped = df_sap_clean.groupby(['PISTA', 'LOTE_KEY', 'ITEM', 'PRODUCTO', 'LOTE'], as_index=False)['SALDO_SAP'].sum()

                    lista_sup = []
                    sem_num = str(semana_obj).strip()
                    nombres_pestaña = [sem_num, f"SEM {sem_num}", f"SEM{sem_num}", f"SEMANA {sem_num}"]
                    
                    for file in archivos_sup:
                        dict_dfs = pd.read_excel(file, sheet_name=None, header=None, dtype=str)
                        target = next((n for n in dict_dfs.keys() if str(n).upper().strip() in [p.upper() for p in nombres_pestaña]), None)
                        
                        if target:
                            df_raw = dict_dfs[target]
                            h_idx = -1
                            for i in range(min(30, len(df_raw))):
                                row_v = [quitar_tildes(x) for x in df_raw.iloc[i].values if pd.notna(x)]
                                if any("LOTE" in val for val in row_v) and any("SALDO" in val for val in row_v):
                                    h_idx = i; break
                            if h_idx != -1:
                                df_s = df_raw.iloc[h_idx + 1:].copy()
                                df_s.columns = [f"{quitar_tildes(x)}_{idx}" for idx, x in enumerate(df_raw.iloc[h_idx])]
                                c_p = next((c for c in df_s.columns if "PRODUC" in c or "DESCRI" in c), None)
                                c_a = next((c for c in df_s.columns if "ALMAC" in c or "PISTA" in c), None)
                                c_l = next((c for c in df_s.columns if "LOTE" in c and "SALDO" not in c), None)
                                c_v = next((c for c in df_s.columns if "SALDO" in c and "INIC" not in c), None)
                                if all([c_p, c_a, c_l, c_v]):
                                    df_s_c = df_s[[c_p, c_a, c_l, c_v]].copy()
                                    df_s_c.columns = ['PRODUCTO_SUP', 'PISTA', 'LOTE_SUP', 'SALDO_FISICO']
                                    df_s_c['PISTA'] = df_s_c['PISTA'].astype(str).str.strip().str.upper().replace('NAN', None).ffill().bfill()
                                    df_s_c['LOTE_KEY'] = df_s_c['LOTE_SUP'].apply(purificar_lote)
                                    df_s_c['SALDO_FISICO'] = pd.to_numeric(df_s_c['SALDO_FISICO'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                                    lista_sup.append(df_s_c)

                    if lista_sup:
                        df_sup_total = pd.concat(lista_sup, ignore_index=True)
                        st.session_state.df_sup_grouped = df_sup_total.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        st.session_state.semana_actual = semana_obj
                        generar_cruce()
                        st.session_state.arqueo_procesado = True
                    else:
                        st.error("❌ No se encontraron datos válidos.")

            except Exception as e:
                st.error(f"Error: {e}")
                
    if st.session_state.arqueo_procesado:
        tab1, tab2, tab3 = st.tabs(["⚠️ Discrepancias y Notas", "🛠️ Conciliador Inteligente", "📋 Inventario Completo"])
        
        with tab1:
            st.subheader("Registros con Diferencias (Limpios de 0s)")
            df_err = st.session_state.cruce_final[st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA"].copy()
            
            if df_err.empty:
                st.success("✅ ¡Inventario cuadrado!")
            else:
                edited_df = st.data_editor(
                    df_err.drop(columns=['LOTE_KEY']),
                    use_container_width=True,
                    hide_index=True,
                    disabled=["PISTA", "ITEM", "PRODUCTO", "LOTE", "SALDO_SAP", "SALDO_FISICO", "DIFERENCIA", "ESTADO"],
                    column_config={"OBSERVACIONES": st.column_config.TextColumn("📝 OBSERVACIONES (Editable)", width="large")}
                )
                
                for _, row in edited_df.iterrows():
                    key = f"{row['PISTA']}_{purificar_lote(row['LOTE'])}"
                    st.session_state.observaciones_memoria[key] = row['OBSERVACIONES']
                    idx_m = st.session_state.cruce_final[(st.session_state.cruce_final['PISTA'] == row['PISTA']) & (st.session_state.cruce_final['LOTE_KEY'] == purificar_lote(row['LOTE']))].index
                    if not idx_m.empty:
                        st.session_state.cruce_final.at[idx_m[0], 'OBSERVACIONES'] = row['OBSERVACIONES']

        with tab2:
            st.markdown("### 🛠️ Fusión de Lotes y Nombres Mal Escritos")
            err_fantasmas = st.session_state.cruce_final[(st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA") & (st.session_state.cruce_final['SALDO_SAP'] == 0) & (st.session_state.cruce_final['SALDO_FISICO'] > 0)]
            
            if err_fantasmas.empty:
                st.success("✅ No hay lotes fantasmas pendientes.")
            else:
                opciones = err_fantasmas.apply(lambda x: f"{x['PISTA']} | Prod: {x['PRODUCTO']} | Lote Físico: {x['LOTE']}", axis=1).tolist()
                sel = st.selectbox("1️⃣ Seleccione el error de digitación del supervisor:", opciones)
                
                if sel:
                    idx_s = opciones.index(sel)
                    row_s = err_fantasmas.iloc[idx_s]
                    
                    df_sap_pista = st.session_state.df_sap_raw[st.session_state.df_sap_raw['PISTA'] == row_s['PISTA']]
                    df_exact = df_sap_pista[df_sap_pista['PRODUCTO'] == row_s['PRODUCTO']]
                    
                    if not df_exact.empty:
                        lotes_validos = df_exact.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()
                        lote_ok_str = st.selectbox(f"2️⃣ Lotes Oficiales de SAP para {row_s['PRODUCTO']}:", sorted(lotes_validos))
                    else:
                        st.warning(f"⚠️ El nombre '{row_s['PRODUCTO']}' tiene un error de escritura. Seleccione el producto correcto de esta lista general:")
                        lotes_validos = df_sap_pista.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()
                        lote_ok_str = st.selectbox(f"2️⃣ Arsenal completo de SAP para la pista {row_s['PISTA']}:", sorted(lotes_validos))
                    
                    if st.button("⚡ FUSIONAR Y JUSTIFICAR", type="primary"):
                        prod_sap_oficial = lote_ok_str.split(" | Lote: ")[0].strip()
                        lote_sap_oficial = lote_ok_str.split(" | Lote: ")[1].strip()
                        
                        mask = (st.session_state.df_sup_grouped['PISTA'] == row_s['PISTA']) & (st.session_state.df_sup_grouped['LOTE_KEY'] == row_s['LOTE_KEY'])
                        
                        key_final = f"{row_s['PISTA']}_{purificar_lote(lote_sap_oficial)}"
                        st.session_state.observaciones_memoria[key_final] = f"Corrección: Nombre/Lote Físico ({row_s['PRODUCTO']} - {row_s['LOTE']}) unificado con SAP ({prod_sap_oficial} - {lote_sap_oficial})"
                        
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_SUP'] = lote_sap_oficial
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_KEY'] = purificar_lote(lote_sap_oficial)
                        st.session_state.df_sup_grouped.loc[mask, 'PRODUCTO_SUP'] = prod_sap_oficial
                        
                        st.session_state.df_sup_grouped = st.session_state.df_sup_grouped.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        
                        generar_cruce()
                        st.rerun()

        with tab3:
            st.subheader("Inventario Consolidado (Libre de Ceros)")
            st.dataframe(st.session_state.cruce_final.drop(columns=['LOTE_KEY']).style.map(
                lambda x: 'background-color: #d4edda; color: #155724' if x == "✅ OK" else '', subset=['ESTADO']
            ), use_container_width=True, hide_index=True)

        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            f_df = st.session_state.cruce_final.drop(columns=['LOTE_KEY'])
            f_df[f_df['ESTADO'] == "❌ DISCREPANCIA"].to_excel(writer, index=False, sheet_name='Diferencias')
            f_df.to_excel(writer, index=False, sheet_name='Total')
            
            for sheetname in writer.sheets:
                worksheet = writer.sheets[sheetname]
                worksheet.auto_filter.ref = worksheet.dimensions 
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 40)
                    worksheet.column_dimensions[column].width = adjusted_width

        st.download_button("📥 Descargar Reporte Ejecutivo", buffer.getvalue(), f"Arqueo_Ejecutivo_Semana_{st.session_state.semana_actual}.xlsx")
