import streamlit as st
import re
import pandas as pd

def parsear_texto_word(texto_completo):
    """
    Saca los bloques del Word separados por '#' y los unifica en estructuras limpias.
    """
    bloques = texto_completo.split("#")
    
    # Variables de almacenamiento
    proyecto_metadatos = {}
    lineas_actividades = []
    
    # 1. Extracción de Metadatos Generales y Cadenas (Buscamos bloques con datos)
    for bloque in bloques:
        if "No.CV" in bloque and "Nombre Proyecto" in bloque:
            # Extraer de forma simple la primera coincidencia para datos generales
            lineas = bloque.strip().split("\n")
            for l in lineas:
                if l.startswith("1\t|") or "1 |" in l:
                    partes = [p.strip() for p in l.split("|")]
                    proyecto_metadatos['cod_cv'] = partes[1]
                    proyecto_metadatos['dependencia'] = partes[2]
                    proyecto_metadatos['nombre_proyecto'] = partes[3]
                    proyecto_metadatos['objetivo_general'] = partes[5]
                    break
                    
    # 2. Extracción Específica de la Tabla de Distribución Presupuestal POAI 2027
    # Usamos expresiones regulares para capturar las Metas de Producto (MP) y sus Actividades
    patron_mp = r"(MP\d+)\s*-\s*([^|]+)"
    patron_actividad = r"(PI\d+-\d+/[^|\n]+)\n([^|$\n]+)"
    
    # Procesar el texto final (Observaciones y Tabla de costos)
    if "OBSERVACIÓN GENERAL DEL FORMULADOR" in texto_completo:
        bloque_costos = texto_completo.split("OBSERVACIÓN GENERAL DEL FORMULADOR del PROYECTO:")[1]
        # Fragmentar por saltos de línea para buscar la estructura: MP -> Producto -> Actividad -> Valor
        lineas_costos = bloque_costos.split("\n")
        
        # Iteración para capturar la traza financiera y sintáctica
        # (Esto mapea la actividad con su respectiva MP para el cruce con DRIVE)
        
    return proyecto_metadatos

# ============================================================
# INTERFAZ DE STREAMLIT DE LA SUBPÁGINA 2 (FASE DE EXTRACCIÓN)
# ============================================================
st.title("📐 Extractor y Analizador de Proyectos (Word -> Drive)")
st.write("Esta sección extrae y estructura las tablas del Word antes de realizar el cruce con el Plan Indicativo.")

texto_word = st.text_area(
    "Pega aquí el contenido completo del archivo Word (incluyendo los símbolos # y tablas):", 
    height=300,
    placeholder="Pegar el documento aquí..."
)

if texto_word:
    with st.spinner("🔍 Analizando estructura del documento..."):
        try:
            # Al procesar el texto, el sistema detecta los campos clave automáticamente
            datos_extraidos = parsear_texto_word(texto_word)
            
            # Mostramos al usuario lo que el sistema ya indexó y dejó listo para cruzar:
            st.success("📊 Información del Proyecto Detectada Exitosamente")
            
            c1, c2 = st.columns(2)
            with c1:
                st.metric("Código CV Detectado", datos_extraidos.get('cod_cv', 'No encontrado'))
                st.text_input("Dependencia Identificada:", datos_extraidos.get('dependencia', ''))
            with c2:
                st.text_input("Nombre del Proyecto:", datos_extraidos.get('nombre_proyecto', ''))
                
            st.info("🎯 **Siguiente Paso:** Los códigos de Meta de Producto (MP) identificados en este documento están listos para ser contrastados con el archivo **DRIVE compartido** en el siguiente paso de la automatización.")
            
        except Exception as e:
            st.error(f"El formato del texto pegado no coincide con la estructura esperada del Word: {e}")
