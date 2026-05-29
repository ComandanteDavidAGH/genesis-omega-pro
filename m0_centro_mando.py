import streamlit as st

def renderizar():
    st.markdown("<h1 class='titulo-principal'>Centro de Mando Omega Pro</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div class='tarjeta-info'>
        <h3>Bienvenido Comandante al Sistema Unificado:</h3>
        <p>Seleccione en el menú lateral la operación que desea realizar hoy. Los módulos están protegidos y operan de forma independiente.</p>
        <ol>
            <li><b>Mantenimiento:</b> Purifique y suba la Sábana SAP a la Bóveda (Plantilla).</li>
            <li><b>Facturación:</b> Cargue la sábana de SAP y los pedidos. Luego valide y facture en el módulo 3.</li>
            <li><b>Ingreso Manual Acelerado:</b> Digite los datos base de sus OS y el sistema calculará e inyectará el resto.</li>
            <li><b>Sincronización:</b> Actualice precios semanalmente simulando la Macro de VBA.</li>
            <li><b>Dominicales:</b> Rastree fechas de operación y recargos con inyección directa.</li>
            <li><b>Arqueo:</b> Auditoría total de pistas contra saldos SAP, con conciliación inteligente.</li>
            <li><b>Radar Hectáreas:</b> Visor dinámico semana a semana y mes a mes para gerencia.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
