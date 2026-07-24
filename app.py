# ⚡ FUNCIÓN DE ORDENAMIENTO CRONOLÓGICO FÍSICO PARA DRIVE Y SUPABASE
def sincronizar_y_ordenar_tabla1_a_supabase():
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        return
    import math
    def sanitizar_valor(val):
        if pd.isna(val): return ""
        if isinstance(val, (float, int)):
            if math.isnan(val) or math.isinf(val): return ""
            return float(val)
        return str(val).strip()

    try:
        supabase = st.session_state['supabase']
        gc = conectar_satelite()
        if not gc: return

        with st.spinner("🔄 Rescatando datos de Drive y restaurando Supabase..."):
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            # Doble Motor: Fórmulas (Drive) y Valores (Supabase)
            t1_formulas = ws_t1.get_all_values(value_render_option='FORMULA')
            t1_valores = ws_t1.get_all_values(value_render_option='UNFORMATTED_VALUE')
            
            idx_header = 4
            for i in range(min(10, len(t1_formulas))):
                if "FINCA" in [str(x).upper().strip() for x in t1_formulas[i]]:
                    idx_header = i
                    break
                    
            # 💥 CORRECCIÓN VITAL: Generar nombres de columna como Supabase los espera (col_0, col_1...)
            num_cols = len(t1_formulas[idx_header])
            cols_db = [f"col_{k}" for k in range(num_cols)]
            
            datos_form = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_formulas[idx_header+1:]]
            datos_val = [r[:num_cols] + [""] * (num_cols - len(r[:num_cols])) for r in t1_valores[idx_header+1:]]
            
            df_form = pd.DataFrame(datos_form, columns=cols_db)
            df_val = pd.DataFrame(datos_val, columns=cols_db)
            
            filas_validas = df_val['col_0'].astype(str).str.strip() != ""
            df_form = df_form[filas_validas].copy()
            df_val = df_val[filas_validas].copy()

            if len(df_val.columns) > 7:
                df_val['fecha_dt'] = pd.to_datetime(df_val['col_7'].astype(str).str.replace("'", "").str.strip(), format='%d/%m/%Y', errors='coerce')
                
                df_val_sorted = df_val.sort_values(by='fecha_dt', ascending=False, na_position='last')
                indices_ordenados = df_val_sorted.index
                
                df_form_sorted = df_form.loc[indices_ordenados]
                df_val_sorted = df_val_sorted.drop(columns=['fecha_dt'])
            else:
                df_val_sorted = df_val
                df_form_sorted = df_form

            # ========================================================
            # 💾 1. RESTAURAR SUPABASE (Con valores limpios 'col_0')
            # ========================================================
            registros_supa = []
            for _, row in df_val_sorted.iterrows():
                fila_limpia = {k: sanitizar_valor(v) for k, v in row.items()}
                registros_supa.append(fila_limpia)

            if registros_supa:
                try:
                    supabase.table("TABLA_1").delete().neq("col_0", "VACIO_FORZADO").execute()
                    tamano_bloque = 200
                    for i in range(0, len(registros_supa), tamano_bloque):
                        supabase.table("TABLA_1").insert(registros_supa[i:i + tamano_bloque]).execute()
                    st.toast("⚡ Supabase restaurado exitosamente.", icon="✅")
                except Exception as e_supa:
                    st.error(f"🚨 Supabase rechazó los datos: {e_supa}")
                    return

            # ========================================================
            # 📝 2. ACTUALIZAR GOOGLE DRIVE (Con fórmulas intactas)
            # ========================================================
            valores_ordenados_drive = df_form_sorted.fillna("").values.tolist()
            if valores_ordenados_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                ws_t1.batch_clear([rango_borrar])
                ws_t1.update(range_name=rango_inicio, values=valores_ordenados_drive, value_input_option='USER_ENTERED')
                
            st.toast("⚡ Google Drive Ordenado.", icon="✅")

    except Exception as e:
        st.error(f"🚨 Error crítico en sincronización: {e}")
