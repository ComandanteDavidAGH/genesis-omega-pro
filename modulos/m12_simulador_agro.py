# --- OBTENER RANGOS PARA FILTROS ---
    # Asignamos fechas por defecto basadas en los datos, pero si falla, usamos un rango amplio
    min_date = df_sim['Fecha_DT'].min().date() if not df_sim['Fecha_DT'].isnull().all() else datetime(2023, 1, 1).date()
    max_date = df_sim['Fecha_DT'].max().date() if not df_sim['Fecha_DT'].isnull().all() else datetime.today().date()
    lista_fincas = sorted(df_sim[col_finca].dropna().unique().tolist())
    opciones_finca = ["🌍 TODAS LAS FINCAS"] + lista_fincas

    # =================================================================
    # 🎛️ PANEL DE CONTROL GERENCIAL (Filtros)
    # =================================================================
    with st.container(border=True):
        st.markdown("#### 🎛️ Filtros de Escenario y Parámetros")
        f1, f2, f3, f4, f5 = st.columns(5)
        
        # 🗓️ Calendarios separados y sin candados de bloqueo
        fecha_ini = f1.date_input("📅 Fecha Inicial", value=min_date)
        fecha_fin = f2.date_input("📆 Fecha Final", value=max_date)
        
        finca_sel = f3.selectbox("📍 Selección de Finca", opciones_finca)
        tarifa_base_hora = f4.number_input("💰 Tarifa Avión (Hora)", value=4606562.0, step=10000.0)
        multiplicador = f5.number_input("✖️ Multiplicador", value=1.112, format="%.3f")

    # --- APLICAR FILTROS DE INTERFAZ ---
    # Usamos las fechas seleccionadas directamente para filtrar
    df_filtrado = df_sim[(df_sim['Fecha_DT'].dt.date >= fecha_ini) & (df_sim['Fecha_DT'].dt.date <= fecha_fin)].copy()

    if finca_sel != "🌍 TODAS LAS FINCAS":
        df_filtrado = df_filtrado[df_filtrado[col_finca] == finca_sel]

    if df_filtrado.empty:
        st.warning("📭 No hay vuelos registrados en este rango de fechas o finca seleccionada.")
        return
