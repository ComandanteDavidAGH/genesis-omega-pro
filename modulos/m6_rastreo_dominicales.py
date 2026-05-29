import streamlit as st
import pandas as pd
import gspread
from datetime import datetime

def ejecutar(procesar_fecha_pesada, limpiar_val_dom):
    st.markdown("<h1 class='titulo-principal'>Rastreo e Inyección de Recargos</h1>", unsafe_allow_html=True)
    
    url_ori = st.text_input(
        "🔗 Pegue URL de GÉNESIS_OMEGA_V2_ESTABLE:", 
        placeholder="Pegue aquí el link..."
    )

    if st.button("🚀 RASTREAR E INYECTAR FALTANTES", use_container_width=True):
        if not url_ori or "http" not in url_ori:
            st.error("❌ Pegue una URL válida.")
        else:
            try:
                if "gcp_credentials" in st.secrets:
                    cred_dict = dict(st.secrets["gcp_credentials"])
                    gc = gspread.service_account_from_dict(cred_dict)
                else:
                    gc = gspread.service_account(filename='credenciales.json')
                    
                with st.spinner("Modo Inyección Exacta Activado..."):
                    url_dest = "https://docs.google.com/spreadsheets/d/1FTiKlHo2UF8lWHk4SrFf9oxTUa2Q_n1l5IK9XFoqQaU/edit"
                    
                    sh_dest = gc.open_by_url(url_dest)
                    ws_dest = sh_dest.sheet1
                    datos_dest = ws_dest.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    
                    max_f = datetime(1900, 1, 1)
                    dict_local = {}
                    
                    for i, row in enumerate(datos_dest):
                        row_padded = row + [""] * (5 - len(row)) if len(row) < 5 else row
                        if i + 1 >= 5 and str(row_padded[1]).strip() != "":
                            f_obj = procesar_fecha_pesada(row_padded[3])
                            if f_obj:
                                if f_obj > max_f: max_f = f_obj
                                dict_local[f"{str(row_padded[1]).strip().upper()}|{f_obj.date()}"] = i + 1

                    st.info(f"📅 Radar Destino: Última fecha validada -> {max_f.strftime('%d/%m/%Y')}")

                    sh_ori = gc.open_by_url(url_ori)
                    ws_ori = next((s for s in sh_ori.worksheets() if "TABLA 1" in s.title.upper()), sh_ori.sheet1)
                    
                    st.write("---")
                    st.write(f"👁️ **RAYOS X ACTIVADOS:** Leyendo Archivo: `{sh_ori.title}` | Pestaña: `{ws_ori.title}`")
                    
                    datos_ori = ws_ori.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    dict_nuevos = {}
                    memoria_fecha = None 
                    recargos_encontrados = 0
                    recargos_ignorados = 0
                    
                    for i, row in enumerate(datos_ori):
                        n_fila = i + 1
                        if n_fila < 6: continue
                        
                        row_padded = row + [""] * (25 - len(row)) if len(row) < 25 else row
                        
                        f_leida = procesar_fecha_pesada(row_padded[7])
                        if f_leida: 
                            memoria_fecha = f_leida 
                        
                        surcharge = limpiar_val_dom(row_padded[20])
                        
                        if surcharge > 0:
                            recargos_encontrados += 1
                            f_operacion = f_leida if f_leida else memoria_fecha
                            
                            if f_operacion and f_operacion > max_f:
                                finca = str(row_padded[2]).strip().upper() if row_padded[2] else "SIN FINCA"
                                ha = limpiar_val_dom(row_padded[5])
                                pista = str(row_padded[23]).strip().upper() if row_padded[23] else ""
                                
                                key = f"{finca}|{f_operacion.date()}"
                                
                                if key in dict_nuevos:
                                    dict_nuevos[key]['ha'] += ha
                                    if not dict_nuevos[key]['pista'] and pista: dict_nuevos[key]['pista'] = pista
                                else:
                                    f_formato = f"{['lunes','martes','miércoles','jueves','viernes','sábado','domingo'][f_operacion.weekday()]}, {['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][f_operacion.month-1]} {f_operacion.day}, {f_operacion.year}"
                                    dict_nuevos[key] = {
                                        'finca': finca, 'ha': ha, 'fec': f_formato,
                                        'sur': surcharge, 'pista': pista, 'semana': f_operacion.isocalendar()[1]
                                    }
                            else:
                                recargos_ignorados += 1

                    st.write(f"📊 **MÉTRICAS:** {recargos_encontrados} Recargos totales | {recargos_ignorados} Ignorados por fecha antigua.")
                    st.write("---")

                    if dict_nuevos:
                        prox_fila = len(datos_dest) + 1 
                        filas_nuevas = [[v['finca'], v['ha'], v['fec'], v['sur'], v['pista'], v['semana']] for v in dict_nuevos.values()]
                        ws_dest.update(f'B{prox_fila}', filas_nuevas, value_input_option='USER_ENTERED')
                        st.success(f"🎯 ¡IMPACTO PERFECTO! {len(filas_nuevas)} registros inyectados empezando en la fila {prox_fila}.")
                        st.balloons()
                    else:
                        st.warning("⚠️ El escáner vio los recargos, pero ninguno era posterior a la fecha del radar.")

            except Exception as e:
                st.error(f"🚨 FALLA DE SISTEMA: {type(e).__name__} - {str(e)}")
