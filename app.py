import streamlit as st

# Configuración de página principal
st.set_page_config(
    page_title="Plataforma de Gestión - SODR",
    page_icon="🏢",
    layout="wide"
)

# Ocultar menús de desarrollo
ocultar_elementos_css = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stAppToolbar {visibility: hidden;}
    </style>
"""
st.markdown(ocultar_elementos_css, unsafe_allow_html=True)

# Contenido de la Bienvenida
st.title("🏢 Sistema Integrado de Trámites y Auditoría - SODR")
st.write("Bienvenido a la plataforma de herramientas de la Subdirección de Ordenamiento y Desarrollo Regional.")
st.markdown("---")

st.subheader("💡 Selecciona un trámite para comenzar:")

# Creamos dos columnas en la pantalla de inicio
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Auditoría de Seguimiento EVAPLAN")
    st.write(
        "Consolidación de Plan Indicativo (PI) con Plan de Acción (PA), "
        "generación de reportes PDF/Excel y construcción dinámica de Prompts para el BOT auditor."
    )
    # Ruta limpia sin emojis y relativa explícita
    st.page_link("./pages/1_Auditoria_EVAPLAN.py", label="Ir a Auditoría EVAPLAN", icon="📊", use_container_width=True)

with col2:
    st.markdown("### 📝 POAI 2027")
    st.write(
        "Módulo destinado a la formulación, revisión y cargue del Plan Operativo Anual de Inversiones (POAI) "
        "para la vigencia 2027. *(En desarrollo)*"
    )
    # Ruta limpia sin emojis y relativa explícita
    st.page_link("./pages/2_POAI_2027.py", label="Explorar Módulo POAI", icon="📝", use_container_width=True)
