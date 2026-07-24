def sincronizar_y_ordenar_tabla1_a_supabase():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        return

    import re
    import unicodedata

    def norm_str(s):
        if not s: return ""
        s = str(s).replace('\n', ' ').strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        return re.sub(r'[^a-zA-Z0-9]', '', s).lower()

    def convertir_fecha_excel(val):
        if pd.isna(val) or val is None or str(val).strip() == "": return ""
        val_str = str(val).strip()
        try:
            num = float(val_str)
            if 30000 < num < 60000:
                return pd.to_datetime(num, unit='D', origin='1899-12-30').strftime('%d/%m/%Y')
        except Exception: pass
        return val_str

    def sanitizar(v):
        if pd.isna(v) or v is None: return ""
        if isinstance(v, (float, int)):
            if math.isnan(v) or math.isinf(v): return 0
            return v
        return str(v).strip()

    try:
        supabase = st.session_state['supabase']
        gc = conectar_satelite()
        if not gc: return

        with st.spinner("🔄 Rescatando Sábana con mapeo exacto de columnas..."):
            db_cols = []
            try:
                res_db = supabase.table("TABLA_1").select("*").limit(1).execute()
                if res_db.data and len(res_db.data) > 0:
                    db_cols = list(res_db.data[0].keys())
            except Exception: pass

            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            t1_formulas = ws_t1.get_all_values(value_render_option='FORMULA')
            t1_valores = ws_t1.get_all_values(value_render_option='FORMATTED_VALUE')
            
            idx_header = 4
            for i in range(min(10, len(t1_formulas))):
                if "FINCA" in [str(x).upper().strip() for x in t1_formulas[i]]:
                    idx_header = i
                    break
                    
            headers_excel = [str(x).replace('\n', ' ').strip() for x in t1_valores[idx_header]]
            num_cols = len(headers_excel)
            
            datos_form = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_formulas[idx_header+1:]]
            datos_val = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_valores[idx_header+1:]]
            
            df_form = pd.DataFrame(datos_form, columns=headers_excel)
            df_val = pd.DataFrame(datos_val, columns=headers_excel)
            
            col_id = headers_excel[0]
            filas_validas = df_val[col_id].astype(str).str.strip() != ""
            df_form = df_form[filas_validas].copy()
            df_val = df_val[filas_validas].copy()

            col_map = {}
            if db_cols:
                db_norm_map = {norm_str(c): c for c in db_cols}
                for h_ex in headers_excel:
                    h_norm = norm_str(h_ex)
                    if h_norm in db_norm_map:
                        col_map[h_ex] = db_norm_map[h_norm]
                    else:
                        col_map[h_ex] = h_ex
            else:
                col_map = {h: h for h in headers_excel}

            col_fecha = next((c for c in headers_excel if "FECHA" in c.upper()), None)

            if col_fecha:
                df_val[col_fecha] = df_val[col_fecha].apply(convertir_fecha_excel)
                df_val['fecha_dt'] = pd.to_datetime(df_val[col_fecha], format='%d/%m/%Y', errors='coerce')
                
                df_val_sorted = df_val.sort_values(by='fecha_dt', ascending=False, na_position='last')
                indices_ord = df_val_sorted.index
                df_form_sorted = df_form.loc[indices_ord]
                df_val_sorted = df_val_sorted.drop(columns=['fecha_dt'])
            else:
                df_val_sorted = df_val
                df_form_sorted = df_form

            registros = []
            for _, row in df_val_sorted.iterrows():
                rec = {}
                for h_ex in headers_excel:
                    if h_ex != "":
                        db_key = col_map.get(h_ex, h_ex)
                        rec[db_key] = sanitizar(row[h_ex])
                registros.append(rec)

            if registros:
                col_id_db = col_map.get(col_id, col_id)
                supabase.table("TABLA_1").delete().neq(col_id_db, "_VACIO_IMPOSIBLE_999_").execute()
                tamano_bloque = 250
                for i in range(0, len(registros), tamano_bloque):
                    supabase.table("TABLA_1").insert(registros[i:i + tamano_bloque]).execute()
                st.toast(f"⚡ Supabase sincronizado ({len(registros)} filas sin NULLs).", icon="✅")

            valores_drive = df_form_sorted[headers_excel].fillna("").values.tolist()
            if valores_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                ws_t1.batch_clear([rango_borrar])
                ws_t1.update(range_name=rango_inicio, values=valores_drive, value_input_option='USER_ENTERED')
                st.toast("⚡ Google Drive Físicamente Ordenado.", icon="✅")

    except Exception as e:
        st.toast(f"🚨 Sincronización en proceso: {e}")
