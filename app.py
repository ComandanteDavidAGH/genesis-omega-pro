# ⚡ FUNCIÓN DE ORDENAMIENTO Y RESTAURACIÓN BLINDADA (CON ESCUDO ANTI-BORRADO)
def sincronizar_y_ordenar_tabla1_a_supabase():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        return

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

        with st.spinner("🔄 Rescatando Sábana de Drive y validando campos..."):
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            t1_formulas = ws_t1.get_all_values(value_render_option='FORMULA')
            t1_valores = ws_t1.get_all_values(value_render_option='UNFORMATTED_VALUE')
            
            idx_header = 4
            for i in range(min(10, len(t1_formulas))):
                if "FINCA" in [str(x).upper().strip() for x in t1_formulas[i]]:
                    idx_header = i
                    break
                    
            headers_real = [str(x).strip() for x in t1_valores[idx_header]]
            num_cols = len(headers_real)
            
            datos_form = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_formulas[idx_header+1:]]
            datos_val = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_valores[idx_header+1:]]
            
            df_form = pd.DataFrame(datos_form, columns=headers_real)
            df_val = pd.DataFrame(datos_val, columns=headers_real)
            
            col_id = headers_real[0]
            filas_validas = df_val[col_id].astype(str).str.strip() != ""
            df_form = df_form[filas_validas].copy()
            df_val = df_val[filas_validas].copy()

            col_fecha = next((c for c in headers_real if "FECHA" in c.upper()), None)

            if col_fecha:
                df_val['fecha_dt'] = pd.to_datetime(df_val[col_fecha].astype(str).str.replace("'", "").str.strip(), format='%d/%m/%Y', errors='coerce')
                df_val_sorted = df_val.sort_values(by='fecha_dt', ascending=False, na_position='last')
                indices_ord = df_val_sorted.index
                df_form_sorted = df_form.loc[indices_ord]
                df_val_sorted = df_val_sorted.drop(columns=['fecha_dt'])
            else:
                df_val_sorted = df_val
                df_form_sorted = df_form

            # Mapeo exacto de los 34 campos
            registros = []
            for _, row in df_val_sorted.iterrows():
                rec = {c: sanitizar(row[c]) for c in headers_real if c != ""}
                registros.append(rec)

            if registros:
                # 💥 ESCUDO BLINDADO: Probar 1 registro antes de borrar nada
                try:
                    test_item = registros[0]
                    supabase.table("TABLA_1").insert([test_item]).execute()
                except Exception as e_test:
                    st.error(f"🚨 Error en estructura de campos de Supabase: {e_test}")
                    return # Aborta inmediatamente SIN borrar nada
                
                # Si la prueba pasó, ejecutamos la sincronización completa
                supabase.table("TABLA_1").delete().neq(col_id, "_VACIO_IMPOSIBLE_999_").execute()
                
                tamano_bloque = 250
                for i in range(0, len(registros), tamano_bloque):
                    supabase.table("TABLA_1").insert(registros[i:i + tamano_bloque]).execute()
                    
                st.toast(f"⚡ Supabase poblado con éxito ({len(registros)} filas).", icon="✅")

            # Actualización física de Google Drive ordenado
            valores_drive = df_form_sorted[headers_real].fillna("").values.tolist()
            if valores_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                ws_t1.batch_clear([rango_borrar])
                ws_t1.update(range_name=rango_inicio, values=valores_drive, value_input_option='USER_ENTERED')
                st.toast("⚡ Google Drive Físicamente Ordenado.", icon="✅")

    except Exception as e:
        st.toast(f"🚨 Sincronización en proceso: {e}")
