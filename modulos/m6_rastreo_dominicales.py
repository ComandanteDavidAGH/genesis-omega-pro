import streamlit as st
import pandas as pd
import gspread
import re
from datetime import datetime

# =================================================================
# ⚡ MOTORES DE CONEXIÓN Y ACCESO SATELITAL (ALTA VELOCIDAD)
# =================================================================

@st.cache_resource(show_spinner=False)
def inicializar_cliente_gspread():
    """ Centraliza la autenticación con Google Cloud una sola vez en RAM """
    try:
        if "gcp_credentials" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_credentials"]))
        return gspread.service_account(filename='credenciales.json')
    except:
        return None

# =================================================================
# 👑 PROCESAMIENTO Y RADAR DE INYECCIÓN DE RECARGOS
# =================================================================

def ejecutar(procesar_fecha_pesada, limpiar_val_dom):
    # Inyección de la línea estética VIP Corporativa de Génesis
    st.markdown("""
    <style>
    .titulo-principal { 
        color: #0d1b2a; 
        border-bottom: 3px solid #d4af37; 
        padding-bottom: 5px; 
        font-family: 'Arial Black', sans-serif; 
    }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { 
        border: 3px solid #0d1b2a !important; 
        border-radius: 8px !important; 
        overflow: hidden !important; 
    }
    
    /* HUD HUD Analítico de Recargos */
    .hud-recargos {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a365d 100%);
        border-left: 5px solid #d4af37; padding: 15px; border-radius: 8px; color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15); margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .hud-recargos-item { text-align: center; flex: 1; }
    .hud-recargos-title { font-size: 11px; font-weight: bold; color: #d4af37; text-transform: uppercase; margin:0; letter-spacing: 1px; }
    .hud-recargos-value { font-size: 20px; font-family: 'Arial Black'; margin: 5px 0 0 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 class='titulo-principal'>Rastreo e Inyección de Recargos</h1>", unsafe_allow_html=True)
    
    url_ori = st.text_input(
        "🔗 Pegue URL de GÉNESIS_OMEGA_V2_ESTABLE:", 
        placeholder="Pegue aquí el link del archivo origen..."
    )

    # Inicialización del cliente gspread desde la memoria caché acelerada
    gc = inicializar_cliente_gspread()
    if gc is None:
        st.error("🚨 Enlace satelital roto con Google Cloud. Verifique sus credenciales.")
        return

    if st.button("🚀 RASTREAR E INYECTAR FALTANTES", use_container_width=True):
        if not url_ori or "http" not in url_ori:
            st.error("❌ Por favor, introduzca una URL válida de Google Sheets.")
        else:
            try:
                with st.spinner("Modo Inyección Exacta Activado..."):
                    url_dest = "https://docs.google.com/spreadsheets/d/1FTiKlHo2UF8lWHk4SrFf9oxTUa2Q_n1l5IK9XFoqQaU/edit"
                    
                    sh_dest = gc.open_by_url(url_dest)
                    ws_dest = sh_dest.sheet1
                    datos_dest = ws_dest.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    
                    max_f = datetime(1900, 1, 1)
                    dict_local = {}
                    
                    # Escáner y mapeo del Radar Destino
                    for i, row in enumerate(datos_dest):
                        row_padded = row + [""] * (5 - len(row)) if len(row) < 5 else row
                        if i + 1 >= 5 and str(row_padded[1]).strip() != "":
                            f_obj = procesar_fecha_pesada(row_padded[3])
                            if f_obj:
                                if f_obj > max_f: max_f = f_obj
                                dict_local[f"{str(row_padded[1]).strip().upper()}|{f_obj.date()}"] = i + 1

                    st.info(f"📅 Radar Destino: Última fecha validada en bóveda -> {max_f.strftime('%d/%m/%Y')}")

                    # Apertura y Rayos X del archivo origen
                    sh_ori = gc.open_by_url(url_ori)
                    ws_ori = next((s for s in sh_ori.worksheets() if "TABLA 1" in s.title.upper()), sh_ori.sheet1)
                    
                    st.write("---")
                    st.write(f"👁️ **RAYOS X ACTIVADOS:** Leyendo Archivo: `{sh_ori.title}` | Pestaña: `{ws_ori.title}`")
                    
                    datos_ori = ws_ori.get_all_values(value_render_option='UNFORMATTED_VALUE')
                    dict_nuevos = {}
                    memoria_fecha = None 
                    recargos_encontrados = 0
                    recargos_ignorados = 0
                    
                    # Listas estáticas de traducción optimizadas fuera del bucle de memoria
                    DIAS_SEMANA = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
                    MESES_ANIO = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

                    # Barrido del archivo de origen
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
                                    if not dict_nuevos[key]['pista'] and pista: 
                                        dict_nuevos[key]['pista'] = pista
                                else:
                                    # Formateo de fecha de alta fidelidad en español latino
                                    f_formato = f"{DIAS_SEMANA[f_operacion.weekday()]}, {MESES_ANIO[f_operacion.month-1]} {f_operacion.day}, {f_operacion.year}"
                                    dict_nuevos[key] = {
                                        'finca': finca, 'ha': ha, 'fec': f_formato,
                                        'sur': surcharge, 'pista': pista, 'semana': f_operacion.isocalendar()[1]
                                    }
                            else:
                                recargos_ignorados += 1

                    # 🚀 DESPLIEGUE DEL HUD DE CONTROL DE RECARGOS
                    st.markdown(f"""
                    <div class="hud-recargos">
                        <div class="hud-recargos-item">
                            <p class="hud-recargos-title">Recargos Encontrados</p>
                            <p class="hud-recargos-value">🧪 {recargos_encontrados} Totales</p>
                        </div>
                        <div class="hud-recargos-item">
                            <p class="hud-recargos-title">Filtro de Antigüedad</p>
                            <p class="hud-recargos-value" style="color: #ff3333;">⚠️ {recargos_ignorados} Ignorados</p>
                        </div>
                        <div class="hud-recargos-item">
                            <p class="hud-recargos-title">Nuevos por Inyectar</p>
                            <p class="hud-recargos-value" style="color: #00ff66;">🚀 {len(dict_nuevos)} Registros</p>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    if dict_nuevos:
                        prox_fila = len(datos_dest) + 1 
                        filas_nuevas = [[v['finca'], v['ha'], v['fec'], v['sur'], v['pista'], v['semana']] for v in dict_nuevos.values()]
                        
                        # Volcado rápido por lote hacia la base destino
                        ws_dest.update(range_name=f'B{prox_fila}', values=filas_nuevas, value_input_option='USER_ENTERED')
                        st.success(f"🎯 ¡IMPACTO PERFECTO! Se inyectaron exitosamente {len(filas_nuevas)} registros nuevos empezando en la fila {prox_fila}.")
                        st.balloons()
                    else:
                        st.warning("⚠️ El escáner detectó recargos en el archivo origen, pero ninguno es posterior al radar de fecha de la base destino.")

            except Exception as e:
                st.error(f"🚨 FALLA CRÍTICA EN EL SISTEMA DE RASTREO: {type(e).__name__} - {str(e)}")
