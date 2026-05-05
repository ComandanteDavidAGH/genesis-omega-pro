import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io

# --- 1. CONFIGURACIÓN DEL NÚCLEO ---
st.set_page_config(page_title="Génesis Omega Pro | AgroAéreo", layout="wide", page_icon="🚀", initial_sidebar_state="expanded")

# --- 2. ARTILLERÍA VISUAL ---
arsenal_css = """
<style>
[data-testid="stToolbarActions"] { display: none !important; }
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #0d1b2a !important; border-right: 4px solid #d4af37; }
[data-testid="stSidebar"] * { color: white !important; font-weight: bold; }
.titulo-principal { color: #0d1b2a; font-family: 'Arial Black', sans-serif; border-bottom: 3px solid #d4af37; text-transform: uppercase;}
.tarjeta-info { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid #0d1b2a; margin-bottom: 20px;}
button[kind="primary"] { background-color: #0d1b2a !important; color: #d4af37 !important; border: 2px solid #d4af37 !important; }
</style>
"""
st.markdown(arsenal_css, unsafe_allow_html=True)

# --- 3. MENÚ TÁCTICO ---
with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: #d4af37;'>🚀 GÉNESIS OMEGA</h2>", unsafe_allow_html=True)
    menu = st.radio("🛰️ NAVEGACIÓN:", ["🏠 Centro de Mando", "📥 1. Buzón de Carga", "⚙️ 2. Validación de Misión", "📊 3. Arqueo y Reportes", "🛡️ Configuración"])
    st.info(f"📅 Operación: {datetime.now().strftime('%Y-%m-%d')}")

# --- 4. LÓGICA DE CARGA ---

if menu == "🏠 Centro de Mando":
    st.markdown("<h1 class='titulo-principal'>Centro de Mando</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3>Estrategia de Validación (La Trinidad):</h3>
        <ol>
            <li><b>Sábana SAP:</b> Validamos Lotes y Precios oficiales.</li>
            <li><b>Pedidos SAP:</b> Validamos lo que se DEBÍA hacer (Fincas/Hectáreas).</li>
            <li><b>Informes Pista:</b> Validamos lo que REALMENTE se hizo.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

elif menu == "📥 1. Buzón de Carga":
    st.markdown("<h1 class='titulo-principal'>Zona de Aterrizaje Tripartita</h1>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("### 📁 Sábana SAP\n*(Lotes y Precios)*")
        f_sabana = st.file_uploader("Subir Export SAP", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="sab")
    with c2:
        st.markdown("### 📝 Pedidos SAP\n*(La Planificación)*")
        f_pedidos = st.file_uploader("Subir Pedidos Diarios", type=["xlsx", "xls", "csv", "CSV", "XLSX"], key="ped")
    with c3:
        st.markdown("### 🚁 Informes Pista\n*(La Realidad)*")
        f_pistas = st.file_uploader("Subir Reportes Pista", type=["xlsx", "xls", "csv", "CSV", "XLSX"], accept_multiple_files=True, key="pis")

    if st.button("🚀 INICIAR PROCESAMIENTO MAESTRO", type="primary", use_container_width=True):
        if f_sabana and f_pedidos and f_pistas:
            with st.spinner("Sincronizando los 3 frentes de batalla (Blindaje Anti-Mayúsculas)..."):
                try:
                    # 1. Leer Sábana
                    bytes_sabana = io.BytesIO(f_sabana.getvalue())
                    nom_sab = f_sabana.name.lower()
                    if nom_sab.endswith('.xlsx') or nom_sab.endswith('.xls'):
                        st.session_state['df_sabana'] = pd.read_excel(bytes_sabana)
                    else:
                        st.session_state['df_sabana'] = pd.read_csv(bytes_sabana, sep=None, engine='python')
                    
                    # 2. Leer Pedidos
                    bytes_pedidos = io.BytesIO(f_pedidos.getvalue())
                    nom_ped = f_pedidos.name.lower()
                    if nom_ped.endswith('.xlsx') or nom_ped.endswith('.xls'):
                        st.session_state['df_pedidos'] = pd.read_excel(bytes_pedidos)
                    else:
                        st.session_state['df_pedidos'] = pd.read_csv(bytes_pedidos, sep=None, engine='python')
                    
                    # 3. Leer Pistas
                    lista_pistas = []
                    errores_pistas = []
                    
                    for f in f_pistas:
                        bytes_pista = io.BytesIO(f.getvalue())
                        nom_pis = f.name.lower()
                        
                        try:
                            if nom_pis.endswith('.csv'):
                                df_raw = pd.read_csv(bytes_pista, header=None, sep=None, engine='python', on_bad_lines='skip')
                                m = df_raw.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                if m.any():
                                    df_m = df_raw.iloc[m.idxmax():].copy()
                                    df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                    df_m['ORIGEN'] = f.name
                                    lista_pistas.append(df_m)
                                    
                            elif nom_pis.endswith('.xlsx'):
                                wb = openpyxl.load_workbook(bytes_pista, read_only=True, data_only=True)
                                visibles = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
                                
                                if visibles:
                                    bytes_pandas = io.BytesIO(f.getvalue())
                                    dict_p = pd.read_excel(bytes_pandas, sheet_name=visibles, header=None)
                                    for name, df in dict_p.items():
                                        m = df.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                        if m.any():
                                            df_m = df.iloc[m.idxmax():].copy()
                                            df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                            df_m['ORIGEN'] = f"{f.name} ({name})"
                                            lista_pistas.append(df_m)
                                            
                            else:
                                dict_p = pd.read_excel(bytes_pista, sheet_name=None, header=None)
                                for name, df in dict_p.items():
                                    m = df.astype(str).apply(lambda x: x.str.contains('MEZCLA PREPARADA', case=False, na=False)).any(axis=1)
                                    if m.any():
                                        df_m = df.iloc[m.idxmax():].copy()
                                        df_m = df_m.dropna(axis=1, how='all').dropna(axis=0, how='all')
                                        df_m['ORIGEN'] = f"{f.name} ({name})"
                                        lista_pistas.append(df_m)
                                        
                        except Exception as e_pista:
                            errores_pistas.append(f"{f.name} ({str(e_pista)})")
                    
                    if lista_pistas:
                        st.session_state['df_pistas'] = pd.concat(lista_pistas, ignore_index=True)
                        st.success(f"✅ ¡Trinidad Sincronizada! SAP: {len(st.session_state['df_sabana'])} filas | Pedidos: {len(st.session_state['df_pedidos'])} filas | Pistas: {len(lista_pistas)} bloques.")
                        
                        if errores_pistas:
                            st.warning(f"⚠️ Nota: Algunos archivos fueron saltados por formato ilegible: {', '.join(errores_pistas)}")
                    else:
                        st.error("🚨 No se encontró información válida de 'MEZCLA PREPARADA' en las pistas.")
                        
                except Exception as e:
                    st.error(f"🚨 Error crítico en el ensamblaje principal: {e}")
        else:
            st.error("🚨 Faltan suministros. Suba los 3 frentes.")

elif menu == "⚙️ 2. Validación de Misión":
    st.markdown("<h1 class='titulo-principal'>Validación Cruzada</h1>")
    if 'df_pedidos' in st.session_state:
        st.info("Aquí compararemos: **Informe Pista vs Pedidos SAP** (Hectáreas) y luego **Informe Pista vs Sábana** (Lotes/Precios).")
        # Próximo paso: Lógica de cruce triple
