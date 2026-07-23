# ⚡ FUNCIÓN DE ORDENAMIENTO CRONOLÓGICO FÍSICO PARA DRIVE Y SUPABASE
def sincronizar_y_ordenar_tabla1_a_supabase():
    """
    Lee TABLA 1 de Drive protegiendo sus fórmulas, la ordena cronológicamente,
    actualiza físicamente Google Sheets y sincroniza con Supabase.
    """
    if 'supabase' not in st.session_state or st.session_state['supabase'] is None:
        return

    try:
        supabase = st.session_state['supabase']
        from modulos.m3_validacion_misiones import obtener_cliente_gspread_unificado
        gc = obtener_cliente_gspread_unificado()
        if not gc: return

        with st.spinner("🔄 Ordenando cronológicamente Google Drive y Supabase..."):
            boveda = gc.open_by_url("https://docs.google.com/spreadsheets/d/1gTu6mAec1qJrxAhw7F-Gl3fVcHaIOnmFUJQYFgqARP4/edit")
            ws_t1 = boveda.worksheet("TABLA 1")
            
            # 💥 CRÍTICO: Leer en modo 'FORMULA' para no destruir las matemáticas de Excel
            t1_raw = ws_t1.get_all_values(value_render_option='FORMULA')
            
            # Buscar dónde empiezan los datos reales
            idx_header = 4
            for i in range(min(10, len(t1_raw))):
                if "FINCA" in [str(x).upper().strip() for x in t1_raw[i]]:
                    idx_header = i
                    break
                    
            cols = [f"col_{k}" for k in range(len(t1_raw[idx_header]))]
            
            # Extraer y cuadrar las filas para evitar desbordes en Pandas
            datos_filas = t1_raw[idx_header+1:]
            max_len = len(cols)
            datos_pad = [r + [""] * (max_len - len(r)) for r in datos_filas]
            
            df_t1 = pd.DataFrame(datos_pad, columns=cols)
            df_t1 = df_t1[df_t1['col_0'].astype(str).str.strip() != ""].copy()

            # ⚡ ORDENAR CRONOLÓGICAMENTE EN MEMORIA
            if len(df_t1.columns) > 7:
                df_t1['fecha_dt'] = pd.to_datetime(df_t1['col_7'], format='%d/%m/%Y', errors='coerce')
                # ascending=False = Las más recientes van a la cima de la tabla
                df_t1 = df_t1.sort_values(by='fecha_dt', ascending=False).drop(columns=['fecha_dt'])

            # 1. 💾 ACTUALIZAR SUPABASE
            registros_supa = df_t1.fillna("").to_dict(orient='records')
            if registros_supa:
                supabase.table("sap_tabla_1_maestro").delete().neq("col_0", "VACIO_FORZADO").execute()
                tamano_bloque = 1000
                for i in range(0, len(registros_supa), tamano_bloque):
                    supabase.table("sap_tabla_1_maestro").insert(registros_supa[i:i + tamano_bloque]).execute()

            # 2. 📝 ORDENAR FÍSICAMENTE GOOGLE DRIVE (EXCEL)
            valores_ordenados_drive = df_t1.fillna("").values.tolist()
            if valores_ordenados_drive:
                rango_inicio = f"A{idx_header + 2}"
                rango_borrar = f"A{idx_header + 2}:ZZ{ws_t1.row_count}"
                
                # Limpiamos el rango viejo de datos para que no queden filas fantasma
                ws_t1.batch_clear([rango_borrar])
                # Inyectamos los datos ordenados activando el motor de fórmulas (USER_ENTERED)
                ws_t1.update(range_name=rango_inicio, values=valores_ordenados_drive, value_input_option='USER_ENTERED')
                
            st.toast("⚡ Nube Físicamente Ordenada y Sincronizada.", icon="✅")

    except Exception as e:
        st.toast(f"🚨 Error en sincronización de fondo: {e}")
