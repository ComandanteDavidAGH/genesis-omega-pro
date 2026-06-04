# --- 🎛️ FILTROS TÁCTICOS AVANZADOS ---
    st.markdown("### 🎛️ Filtros de Operación y Tiempo")
    
    # ⚡ AQUÍ ESTÁ EL CAMBIO: 3 columnas en lugar de 2
    t1, t2, t3 = st.columns(3)
    
    años_disp = ["TODOS"] + sorted(df_dash['AÑO'].unique().tolist(), reverse=True)
    año_sel = t1.selectbox("📅 AÑO FISCAL", años_disp, index=0)
    
    trimestres = {"TODOS": 0, "Q1 (Ene-Mar)": 1, "Q2 (Abr-Jun)": 2, "Q3 (Jul-Sep)": 3, "Q4 (Oct-Dic)": 4}
    trim_sel = t2.selectbox("📊 TRIMESTRE", list(trimestres.keys()))

    # ⚡ AQUÍ ESTÁ EL CAMBIO: El nuevo selector de mes
    meses_lista = ["TODOS", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    mes_sel = t3.selectbox("📆 MES", meses_lista)

    f1, f2, f3 = st.columns(3)
    fincas_disp = ["TODAS"] + sorted(df_dash['FINCA'].astype(str).unique().tolist())
    pilotos_disp = ["TODOS"] + sorted(df_dash['PILOTO'].astype(str).unique().tolist())
    hks_disp = ["TODAS"] + sorted(df_dash['HK'].astype(str).unique().tolist())
    
    finca_filtro = f1.selectbox("📍 FINCA", fincas_disp)
    piloto_filtro = f2.selectbox("👨‍✈️ PILOTO", pilotos_disp)
    hk_filtro = f3.selectbox("✈️ MATRÍCULA (HK)", hks_disp)

    # 🎯 FILTRADO EN MILISEGUNDOS DESDE LA RAM LOCAL
    df_filtrado = df_dash.copy()
    if año_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(año_sel)]
    if trimestres[trim_sel] != 0: df_filtrado = df_filtrado[df_filtrado['TRIMESTRE'] == trimestres[trim_sel]]
    
    # ⚡ AQUÍ ESTÁ EL CAMBIO: La lógica que filtra la base de datos según el mes
    if mes_sel != "TODOS": df_filtrado = df_filtrado[df_filtrado['MES_NOMBRE'] == mes_sel]
