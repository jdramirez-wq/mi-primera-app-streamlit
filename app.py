# Creamos dos columnas en la pantalla de inicio
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Auditoría de Seguimiento EVAPLAN")
    st.write(
        "Consolidación de Plan Indicativo (PI) con Plan de Acción (PA), "
        "generación de reportes PDF/Excel y construcción dinámica de Prompts para el BOT auditor."
    )
    # Cambiamos a la ruta relativa explícita "./pages/..."
    st.page_link("./pages/1_Auditoria_EVAPLAN.py", label="Ir a Auditoría EVAPLAN", icon="📊", use_container_width=True)

with col2:
    st.markdown("### 📝 POAI 2027")
    st.write(
        "Módulo destinado a la formulación, revisión y cargue del Plan Operativo Anual de Inversiones (POAI) "
        "para la vigencia 2027. *(En desarrollo)*"
    )
    # Cambiamos a la ruta relativa explícita "./pages/..."
    st.page_link("./pages/2_POAI_2027.py", label="Explorar Módulo POAI", icon="📝", use_container_width=True)
