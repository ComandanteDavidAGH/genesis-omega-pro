# 1. CÁLCULO DE HECTÁREAS (Columna F)
        # AQUÍ SÍ BORRAMOS DUPLICADOS: Para que un vuelo de 50ha con 3 químicos 
        # no se sume como 150ha.
        subset_unicos = ['FECHA_DT', 'FINCA_MAESTRA', 'OS_MAESTRA', 'AREA_NUM']
        df_area_a = df_periodo_a.drop_duplicates(subset=subset_unicos)
        
        area_a = df_area_a['AREA_NUM'].sum() if not df_area_a.empty else 0.0


        # 2. CÁLCULO DEL PROMEDIO GLOBAL (Columna W)
        # AQUÍ NO BORRAMOS NADA: Promediamos TODA la columna con .mean()
        # Esto hace el trabajo EXACTAMENTE igual a la función "Media" de su Excel.
        costo_a = df_periodo_a['COSTO_NUM'].mean() if not df_periodo_a.empty else 0
