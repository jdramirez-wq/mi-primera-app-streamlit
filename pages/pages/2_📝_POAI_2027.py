import streamlit as st

st.set_page_config(page_title="POAI 2027", page_icon="📝", layout="wide")
st.markdown("<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;} .stAppToolbar {visibility: hidden;}</style>", unsafe_allow_html=True)

st.title("📝 Formulación y Revisión POAI 2027")
st.write("Subdirección de Ordenamiento y Desarrollo Regional (SODR)")
st.markdown("---")

st.info("🚧 **Módulo en Construcción:** Este espacio está reservado para la estructuración y validación del Plan Operativo Anual de Inversiones (POAI) de la vigencia 2027.")

# Ejemplo de lo que podrías simular para el futuro
st.subheader("Próximas funcionalidades:")
st.markdown("""
* 📤 Carga de proyectos viabilizados en el Banco de Programas y Proyectos de Inversión Departamental.
* ⚖️ Validación de techos presupuestales por sector administrativo.
* 📋 Consolidación de fichas MGA / EVAPLAN para la vigencia 2027.
""")
