import streamlit as st
import pandas as pd
import io
import streamlit.components.v1 as components

def ejecutar(quitar_tildes, purificar_lote):
    st.markdown("<h1 class='titulo-principal'>⚖️ Arqueo de Inventarios y Conciliación</h1>", unsafe_allow_html=True)
    
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
        cruce['SALDO_SAP'] = cruce['SALDO_SAP'].fillna(0).round(3)
        cruce['SALDO_FISICO'] = cruce['SALDO_FISICO'].fillna(0).round(3)
        cruce = cruce[~((cruce['SALDO_SAP'] == 0) & (cruce['SALDO_FISICO'] == 0))]
        cruce['DIFERENCIA'] = (cruce['SALDO_FISICO'] - cruce['SALDO_SAP']).round(3)
        cruce['ESTADO'] = cruce['DIFERENCIA'].apply(lambda x: "✅ OK" if abs(x) <= 0.05 else "❌ DISCREPANCIA")
        cruce['OBSERVACIONES'] = ""
        for idx, row in cruce.iterrows():
            key = f"{row['PISTA']}_{row['LOTE_KEY']}"
            if key in st.session_state.observaciones_memoria: cruce.at[idx, 'OBSERVACIONES'] = st.session_state.observaciones_memoria[key]
            elif row['SALDO_SAP'] > 0 and row['SALDO_FISICO'] == 0: cruce.at[idx, 'OBSERVACIONES'] = "SUGERIDO: Entrega / Traslado / Pendiente por Facturar"
        st.session_state.cruce_final = cruce[['PISTA', 'ITEM', 'PRODUCTO', 'LOTE_KEY', 'LOTE', 'SALDO_SAP', 'SALDO_FISICO', 'DIFERENCIA', 'ESTADO', 'OBSERVACIONES']].sort_values(by=['PISTA', 'PRODUCTO'])

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 INICIAR ARQUEO ESTRATÉGICO", type="primary", use_container_width=True):
        if not archivo_sap or not archivos_sup or not semana_obj: st.error("❌ Faltan suministros.")
        else:
            try:
                with st.spinner("Desplegando analista de inventarios..."):
                    st.session_state.observaciones_memoria = {}
                    sap_file = archivo_sap[0] if isinstance(archivo_sap, list) else archivo_sap
                    if sap_file.name.lower().endswith('.xlsx') or sap_file.name.lower().endswith('.xls'): df_sap = pd.read_excel(sap_file)
                    else:
                        try: df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='utf-8')
                        except UnicodeDecodeError: sap_file.seek(0); df_sap = pd.read_csv(sap_file, sep=None, engine='python', encoding='latin1')

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
                    df_sap_clean['SALDO_SAP'] = pd.to_numeric(df_sap_clean['SALDO_SAP'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                    st.session_state.df_sap_raw = df_sap_clean 
                    st.session_state.df_sap_grouped = df_sap_clean.groupby(['PISTA', 'LOTE_KEY', 'ITEM', 'PRODUCTO', 'LOTE'], as_index=False)['SALDO_SAP'].sum()

                    lista_sup = []; sem_num = str(semana_obj).strip()
                    nombres_pestaña = [sem_num, f"SEM {sem_num}", f"SEM{sem_num}", f"SEMANA {sem_num}"]
                    for file in archivos_sup:
                        dict_dfs = pd.read_excel(file, sheet_name=None, header=None, dtype=str)
                        target = next((n for n in dict_dfs.keys() if str(n).upper().strip() in [p.upper() for p in nombres_pestaña]), None)
                        if target:
                            df_raw = dict_dfs[target]; h_idx = -1
                            for i in range(min(30, len(df_raw))):
                                row_v = [quitar_tildes(x) for x in df_raw.iloc[i].values if pd.notna(x)]
                                if any("LOTE" in val for val in row_v) and any("SALDO" in val for val in row_v): h_idx = i; break
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
                        st.session_state.df_sup_grouped = pd.concat(lista_sup, ignore_index=True).groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        st.session_state.semana_actual = semana_obj
                        generar_cruce()
                        st.session_state.arqueo_procesado = True
                    else: st.error("❌ No se encontraron datos válidos.")
            except Exception as e: st.error(f"🚨 Error: {e}")
                
    if st.session_state.arqueo_procesado:
        tab1, tab2, tab3 = st.tabs(["⚠️ Discrepancias", "🛠️ Conciliador", "📋 Inventario Completo"])
        
        with tab1:
            df_err = st.session_state.cruce_final[st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA"].copy()
            if df_err.empty: st.success("✅ ¡Inventario cuadrado!")
            else:
                edited_df = st.data_editor(df_err.drop(columns=['LOTE_KEY']), use_container_width=True, hide_index=True, disabled=["PISTA", "ITEM", "PRODUCTO", "LOTE", "SALDO_SAP", "SALDO_FISICO", "DIFERENCIA", "ESTADO"], column_config={"SALDO_SAP": st.column_config.NumberColumn("SALDO SAP", format="%.3f"), "SALDO_FISICO": st.column_config.NumberColumn("SALDO FÍSICO", format="%.3f"), "DIFERENCIA": st.column_config.NumberColumn("DIFERENCIA", format="%.3f"), "OBSERVACIONES": st.column_config.TextColumn("📝 OBSERVACIONES (Editable)", width="large")})
                for _, row in edited_df.iterrows():
                    key = f"{row['PISTA']}_{purificar_lote(row['LOTE'])}"
                    st.session_state.observaciones_memoria[key] = row['OBSERVACIONES']
                    idx_m = st.session_state.cruce_final[(st.session_state.cruce_final['PISTA'] == row['PISTA']) & (st.session_state.cruce_final['LOTE_KEY'] == purificar_lote(row['LOTE']))].index
                    if not idx_m.empty: st.session_state.cruce_final.at[idx_m[0], 'OBSERVACIONES'] = row['OBSERVACIONES']

        with tab2:
            err_fantasmas = st.session_state.cruce_final[(st.session_state.cruce_final['ESTADO'] == "❌ DISCREPANCIA") & (st.session_state.cruce_final['SALDO_SAP'] == 0) & (st.session_state.cruce_final['SALDO_FISICO'] > 0)]
            if err_fantasmas.empty: st.success("✅ No hay lotes fantasmas.")
            else:
                opciones = err_fantasmas.apply(lambda x: f"{x['PISTA']} | Prod: {x['PRODUCTO']} | Lote Físico: {x['LOTE']}", axis=1).tolist()
                sel = st.selectbox("1️⃣ Seleccione el error de digitación:", opciones)
                if sel:
                    row_s = err_fantasmas.iloc[opciones.index(sel)]
                    df_sap_pista = st.session_state.df_sap_raw[st.session_state.df_sap_raw['PISTA'] == row_s['PISTA']]
                    df_exact = df_sap_pista[df_sap_pista['PRODUCTO'] == row_s['PRODUCTO']]
                    if not df_exact.empty: lote_ok_str = st.selectbox(f"2️⃣ Lotes Oficiales:", sorted(df_exact.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()))
                    else: lote_ok_str = st.selectbox(f"2️⃣ Arsenal completo de la pista:", sorted(df_sap_pista.apply(lambda x: f"{x['PRODUCTO']} | Lote: {x['LOTE']}", axis=1).unique().tolist()))
                    
                    if st.button("⚡ FUSIONAR", type="primary"):
                        prod_sap, lote_sap = lote_ok_str.split(" | Lote: ")[0].strip(), lote_ok_str.split(" | Lote: ")[1].strip()
                        mask = (st.session_state.df_sup_grouped['PISTA'] == row_s['PISTA']) & (st.session_state.df_sup_grouped['LOTE_KEY'] == row_s['LOTE_KEY'])
                        st.session_state.observaciones_memoria[f"{row_s['PISTA']}_{purificar_lote(lote_sap)}"] = f"Corrección unificada con SAP ({prod_sap} - {lote_sap})"
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_SUP'] = lote_sap
                        st.session_state.df_sup_grouped.loc[mask, 'LOTE_KEY'] = purificar_lote(lote_sap)
                        st.session_state.df_sup_grouped.loc[mask, 'PRODUCTO_SUP'] = prod_sap
                        st.session_state.df_sup_grouped = st.session_state.df_sup_grouped.groupby(['PISTA', 'LOTE_KEY', 'PRODUCTO_SUP', 'LOTE_SUP'], as_index=False)['SALDO_FISICO'].sum()
                        generar_cruce()
                        st.rerun()

        with tab3:
            st.dataframe(st.session_state.cruce_final.drop(columns=['LOTE_KEY']).style.map(lambda x: 'background-color: #d4edda; color: #155724' if x == "✅ OK" else '', subset=['ESTADO']), use_container_width=True, hide_index=True, column_config={"SALDO_SAP": st.column_config.NumberColumn("SALDO SAP", format="%.3f"), "SALDO_FISICO": st.column_config.NumberColumn("SALDO FÍSICO", format="%.3f"), "DIFERENCIA": st.column_config.NumberColumn("DIFERENCIA", format="%.3f")})

        st.markdown("---")
        
        # --- ZONA DUAL: EXCEL Y VISOR HTML PARA PDF (CON DESCARGA DIRECTA INYECTADA) ---
        col_dw1, col_dw2 = st.columns(2)
        
        with col_dw1:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                f_df = st.session_state.cruce_final.drop(columns=['LOTE_KEY'])
                f_df[f_df['ESTADO'] == "❌ DISCREPANCIA"].to_excel(writer, index=False, sheet_name='Diferencias')
                f_df.to_excel(writer, index=False, sheet_name='Total')
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                borde_fino = Border(left=Side(style='thin', color='D1D1D1'), right=Side(style='thin', color='D1D1D1'), top=Side(style='thin', color='D1D1D1'), bottom=Side(style='thin', color='D1D1D1'))
                fondo_navy, texto_blanco = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid"), Font(color="FFFFFF", bold=True)
                for sheetname in writer.sheets:
                    worksheet = writer.sheets[sheetname]
                    worksheet.auto_filter.ref = worksheet.dimensions 
                    for r_idx in range(2, worksheet.max_row + 1):
                        worksheet.cell(row=r_idx, column=7).value = f"=F{r_idx}-E{r_idx}"
                        worksheet.cell(row=r_idx, column=8).value = f'=IF(ABS(G{r_idx})<=0.05, "✅ OK", "❌ DISCREPANCIA")'
                        worksheet.cell(row=r_idx, column=5).number_format = '0.000'; worksheet.cell(row=r_idx, column=6).number_format = '0.000'; worksheet.cell(row=r_idx, column=7).number_format = '0.000'
                    for row_cells in worksheet.iter_rows():
                        for cell in row_cells:
                            cell.border = borde_fino
                            if cell.row == 1: cell.fill = fondo_navy; cell.font = texto_blanco; cell.alignment = Alignment(horizontal='center', vertical='center')
                            elif cell.column in [5, 6, 7]: cell.alignment = Alignment(horizontal='right')
                            elif cell.column == 8: cell.alignment = Alignment(horizontal='center')
                    for col in worksheet.columns: worksheet.column_dimensions[col[0].column_letter].width = min(max(max(len(str(c.value or '')) for c in col) + 4, 12), 42)

            st.download_button("📊 DESCARGAR EXCEL VIVO", buffer.getvalue(), f"Arqueo_Excel_Semana_{st.session_state.semana_actual}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        with col_dw2:
            if st.button("📄 ACTIVAR CENTRO DE EMISIÓN DE PDF", type="primary", use_container_width=True):
                css_vip = """<style>body { font-family: Helvetica, sans-serif; background: white; color: black; font-size: 11px; } .b-print { padding: 20px; } table { width: 100%; border-collapse: collapse; margin-bottom: 20px; } th { background-color: #0d1b2a; color: #d4af37; border: 1px solid #000; padding: 6px; text-align: center; font-size: 12px; } td { border: 1px solid #000; padding: 4px; text-align: center; } .td-left { text-align: left; } .title { font-size: 20px; color: #0d1b2a; font-weight: bold; text-align: center; margin: 0; } .subtitle { font-size: 14px; color: #d4af37; text-align: center; margin: 0 0 20px 0; font-weight: bold; } .firmas-container { display: flex; justify-content: space-around; margin-top: 50px; page-break-inside: avoid; } .firma-box { text-align: center; width: 40%; border-top: 2px solid #0d1b2a; padding-top: 5px; font-weight: bold; color: #0d1b2a; } @media print { @page { size: A4 landscape; margin: 10mm; } body { background: white; -webkit-print-color-adjust: exact; print-color-adjust: exact; } .no-print { display: none !important; } .salto-pagina { page-break-after: always; } }</style>"""
                
                df_agrupado = st.session_state.cruce_final.copy()
                pistas = sorted(df_agrupado['PISTA'].unique())
                
                # 🧪 INYECTOR DEL CDN DE HTML2PDF PARA DESCARGA DIRECTA SIN REQUISITOS EN EL SERVIDOR
                html_masivo = f"""
                <html>
                <head>
                    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
                    <script>
                        function imprimir() {{ window.print(); }}
                        function descargarPDF() {{
                            var element = document.getElementById('contenido-reporte');
                            var opt = {{
                                margin:       [10, 10, 10, 10],
                                filename:     'Reporte_Arqueo_Semana_{st.session_state.semana_actual}.pdf',
                                image:        {{ type: 'jpeg', quality: 0.98 }},
                                html2canvas:  {{ scale: 2, useCORS: true }},
                                jsPDF:        {{ unit: 'mm', format: 'a4', orientation: 'landscape' }},
                                pagebreak:    {{ mode: ['css', 'legacy'] }}
                            }};
                            html2pdf().set(opt).from(element).save();
                        }}
                    </script>
                    {css_vip}
                </head>
                <body>
                    <div class="no-print" style="position: sticky; top: 0; background: white; padding: 10px; z-index: 100; border-bottom: 2px solid #0d1b2a; text-align: right;">
                        <button onclick="descargarPDF()" style="background:#0d1b2a; color:#d4af37; border:2px solid #d4af37; padding:10px 20px; cursor:pointer; border-radius:6px; font-weight:bold; font-family:'Arial Black'; margin-right: 10px;">📥 DESCARGAR PDF DIRECTO</button>
                        <button onclick="imprimir()" style="background:#4a5568; color:white; border:2px solid #4a5568; padding:10px 20px; cursor:pointer; border-radius:6px; font-weight:bold; font-family:'Arial Black';">🖨️ PANEL DE IMPRESIÓN</button>
                    </div>
                    <div id="contenido-reporte">
                """
                
                for i, pista in enumerate(pistas):
                    df_pista = df_agrupado[df_agrupado['PISTA'] == pista]
                    salto = "salto-pagina" if i < len(pistas) - 1 else ""
                    
                    html_masivo += f"""<div class="b-print {salto}">
                        <p class="title">REPORTE OFICIAL DE ARQUEO DE INVENTARIOS</p>
                        <p class="subtitle">BASE OPERATIVA: {pista} | SEMANA: {st.session_state.semana_actual}</p>
                        <table><tr><th style="width:10%;">CÓDIGO</th><th style="width:30%;">PRODUCTO</th><th style="width:15%;">LOTE</th><th style="width:10%;">S. SAP</th><th style="width:10%;">S. FÍSICO</th><th style="width:10%;">DIF.</th><th style="width:15%;">ESTADO</th></tr>"""
                    
                    for _, row in df_pista.iterrows():
                        st_color = "#155724" if "OK" in str(row['ESTADO']) else "#721c24"
                        bg_color = "#d4edda" if "OK" in str(row['ESTADO']) else "#f8d7da"
                        val_dif = f"+{row['DIFERENCIA']:.3f}" if row['DIFERENCIA'] > 0 else f"{row['DIFERENCIA']:.3f}"
                        html_masivo += f"<tr><td>{row['ITEM']}</td><td class='td-left'>{row['PRODUCTO']}</td><td>{row['LOTE']}</td><td>{row['SALDO_SAP']:.3f}</td><td>{row['SALDO_FISICO']:.3f}</td><td style='color:{st_color};'><b>{val_dif}</b></td><td style='color:{st_color}; background-color:{bg_color}; font-weight:bold;'>{row['ESTADO']}</td></tr>"
                    
                    html_masivo += """</table>
                        <div class='firmas-container'>
                            <div class='firma-box'>FIRMA SUPERVISOR DE PISTA</div>
                            <div class='firma-box'>FIRMA AUDITOR DE INVENTARIOS</div>
                        </div></div>"""
                
                html_masivo += "</div></body></html>"
                st.info("💡 **Coordenada Activada:** Use el botón azul de arriba para forzar la descarga del PDF de forma directa a su disco local.")
                components.html(html_masivo, height=600, scrolling=True)
