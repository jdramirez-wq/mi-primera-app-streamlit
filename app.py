# Configuración de página de Streamlit
st.set_page_config(
    page_title="Consolidador PI + PA + PDF",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# OCULTAR MENÚS Y OPCIONES DE DESARROLLO (CÓDIGO CSS)
# ============================================================
ocultar_elementos_css = """
    <style>
    /* Ocultar el botón de opciones en la esquina superior derecha */
    #MainMenu {visibility: hidden;}
    
    /* Ocultar la barra de estado y el footer de Streamlit */
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Alternativa para versiones nuevas de Streamlit que oculta el botón de despliegue */
    .stAppToolbar {visibility: hidden;}
    </style>
"""
st.markdown(ocultar_elementos_css, unsafe_allow_html=True)
