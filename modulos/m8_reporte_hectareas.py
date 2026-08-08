if modo_historico_global:
            st.success("🌐 **MODO MACRO ACTIVADO:** Tienes el control total de la historia operativa. Selecciona el rango de tiempo a analizar.")
            
            # 💥 NUEVO: Selectores de fecha exclusivos para el modo Macro
            cm1, cm2 = st.columns(2)
            fecha_macro_ini = cm1.date_input("F. INICIAL (MACRO):", value=date(2017, 1, 1), min_value=date(2017, 1, 1), max_value=date(2030, 12, 31), key="m8_mac_ini")
            fecha_macro_fin = cm2.date_input("F. FINAL (MACRO):", value=date(2026, 12, 31), min_value=date(2017, 1, 1), max_value=date(2030, 12, 31), key="m8_mac_fin")
            
            # Filtro por fechas
            df_macro = super_base_bi[(super_base_bi['AREA_NUM'] > 0) & 
                                     (super_base_bi['FECHA_DT'].dt.date >= fecha_macro_ini) & 
                                     (super_base_bi['FECHA_DT'].dt.date <= fecha_macro_fin)].copy()
                                     
            if not col_pista: col_pista = "PISTA" # Fallback

            st.markdown("---")
            st.markdown(f"#### 📅 Evolución Anual de Hectáreas por Base")
            
            if df_macro.empty:
                st.warning("⚠️ No hay datos históricos registrados en este rango específico de fechas.")
            else:
                pivot_anual = pd.pivot_table(df_macro, values='AREA_NUM', index='AÑO', columns=col_pista, aggfunc='sum', fill_value=0)
                pivot_anual['TOTAL AÑO'] = pivot_anual.sum(axis=1)
                pivot_anual.loc['TOTAL HISTÓRICO'] = pivot_anual.sum(axis=0)
                st.dataframe(pivot_anual.style.format(fmt_latino).background_gradient(cmap="Blues", axis=None), use_container_width=True)

                df_graf = pivot_anual.drop('TOTAL HISTÓRICO', errors='ignore').drop(columns=['TOTAL AÑO'], errors='ignore')
                fig_macro = px.line(df_graf, markers=True, title="<b>Curva Histórica de Aplicación por Pista</b>")
                fig_macro.update_layout(xaxis_title="Año", yaxis_title="Hectáreas Netas")
                st.plotly_chart(fig_macro, use_container_width=True)

                st.markdown("#### 📆 Desglose Detallado: Año y Mes")
                df_macro = df_macro.sort_values(['AÑO', 'MES'])
                pivot_mes = pd.pivot_table(df_macro, values='AREA_NUM', index=['AÑO', 'MES_NMB'], columns=col_pista, aggfunc='sum', fill_value=0, sort=False)
                pivot_mes['TOTAL MES'] = pivot_mes.sum(axis=1)
                st.dataframe(pivot_mes.style.format(fmt_latino).background_gradient(cmap="YlGn", axis=None), use_container_width=True)

                # 💥 EL HELICÓPTERO DE EXTRACCIÓN VIP (Reporte Macro Formateado)
                st.markdown("---")
                buffer_macro = io.BytesIO()
                
                # Preparamos las tablas planas para el Excel
                df_exp_anual = pivot_anual.reset_index()
                df_exp_mes = pivot_mes.reset_index()
                
                rango_txt_macro = f"{fecha_macro_ini.day}/{fecha_macro_ini.month}/{fecha_macro_ini.year} ⸺ {fecha_macro_fin.day}/{fecha_macro_fin.month}/{fecha_macro_fin.year}"

                with pd.ExcelWriter(buffer_macro, engine='openpyxl') as writer:
                    df_exp_anual.to_excel(writer, sheet_name='Resumen_Anual', index=False, startrow=3)
                    df_exp_mes.to_excel(writer, sheet_name='Desglose_Mensual', index=False, startrow=3)
                    
                    for s_name, df_sheet in [('Resumen_Anual', df_exp_anual), ('Desglose_Mensual', df_exp_mes)]:
                        ws = writer.sheets[s_name]
                        
                        # Título Principal
                        ws['A1'] = f"REPORTE MACRO-HISTÓRICO - {s_name.replace('_', ' ').upper()}"
                        ws['A1'].font = Font(size=14, bold=True, color="FFFFFF")
                        ws['A1'].fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
                        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(df_sheet.columns))
                        
                        # Subtítulo (Fechas dinámicas conectadas al filtro)
                        ws['A3'] = f"Período Analizado: {rango_txt_macro}"
                        ws['A3'].font = Font(italic=True, color="555555", bold=True)
                        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(df_sheet.columns))
                        
                        # Encabezados Dorados
                        header_fill = PatternFill(start_color="D4AF37", end_color="D4AF37", fill_type="solid")
                        header_font = Font(bold=True, color="000000")
                        for col_num in range(1, len(df_sheet.columns) + 1):
                            cell = ws.cell(row=4, column=col_num)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            ws.column_dimensions[get_column_letter(col_num)].width = 16
                        
                        # Formato de números y centrado
                        for r_idx in range(5, len(df_sheet) + 5):
                            for c_idx in range(1, len(df_sheet.columns) + 1):
                                cell = ws.cell(row=r_idx, column=c_idx)
                                if isinstance(cell.value, (int, float)):
                                    cell.number_format = '#,##0.00'
                                cell.alignment = Alignment(horizontal='center')

                st.download_button(
                    label="💾 DESCARGAR HISTÓRICO MACRO EN EXCEL VIP",
                    data=buffer_macro.getvalue(),
                    file_name=f"Reporte_Macro_{fecha_macro_ini.strftime('%Y%m%d')}_{fecha_macro_fin.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
