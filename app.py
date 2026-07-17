import streamlit as st

# Configuración de página principal (Inicia con el menú expandido de forma nativa)
st.set_page_config(
    page_title="Plataforma de Gestión - SODR",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS seguro: SOLO oculta la línea decorativa superior y el pie de página, 
# dejando intactos los botones de navegación y despliegue de Streamlit.
estilo_seguro_css = """
    <style>
    /* Oculta la línea roja/decorativa superior del header */
    div[data-testid="stHeader"] {background-color: transparent;}
    /* Oculta el pie de página de marca */
    footer {visibility: hidden;}
    </style>
"""
st.markdown(estilo_seguro_css, unsafe_allow_html=True)

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
    # Enrutamiento directo y limpio por nombre de archivo
    st.page_link("pages/1_Auditoria_EVAPLAN.py", label="Ir a Auditoría EVAPLAN", icon="📊", use_container_width=True)

with col2:
    st.markdown("### 📝 POAI 2027")
    st.write(
        "Módulo destinado a la formulación, revisión y cargue del Plan Operativo Anual de Inversiones (POAI) "
        "para la vigencia 2027. *(En desarrollo)*"
    )
    # Enrutamiento directo y limpio por nombre de archivo
    st.page_link("pages/2_POAI_2027.py", label="Explorar Módulo POAI", icon="📝", use_container_width=True)
