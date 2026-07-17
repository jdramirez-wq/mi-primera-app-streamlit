import streamlit as st

# Configuración de página de Streamlit para la Subpágina 2
st.set_page_config(
    page_title="POAI 2027",
    page_icon="📝",
    layout="wide"
)

# Estilo CSS seguro: Evita que desaparezca el botón de despliegue (>)
estilo_seguro_p2_css = """
    <style>
    /* Oculta la línea roja decorativa del header */
    div[data-testid="stHeader"] {background-color: transparent;}
    /* Oculta el pie de página */
    footer {visibility: hidden;}
    </style>
"""
st.markdown(estilo_seguro_p2_css, unsafe_allow_html=True)

# Contenido del Módulo
st.title("📝 Formulación y Revisión POAI 2027")
st.write("Herramienta de soporte para el Plan Operativo Anual de Inversiones de la vigencia 2027.")
st.markdown("---")

st.warning("⚠️ Este módulo se encuentra actualmente en fase de codificación y desarrollo técnico.")
