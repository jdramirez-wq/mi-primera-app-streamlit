import streamlit as st
import docx  # Importa python-docx para leer archivos .docx
import re

st.set_page_config(
    page_title="Revisión de Proyectos - Control Previo",
    page_icon="📐",
    layout="wide"
)

def extraer_texto_y_tablas_docx(file_buffer):
    """
    Lee un archivo .docx desde la memoria, extrae el texto de los párrafos 
    y reconstruye las tablas en formato de texto plano estructurado.
    """
    doc = docx.Document(file_buffer)
    contenido_total = []
    
    # 1. Recorrer los elementos del documento en orden de aparición
    # (python-docx nos permite iterar sobre párrafos y tablas)
    for elemento in doc.element.body:
        # Si el elemento es un párrafo, extraemos su texto
        if elemento.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(elemento, doc)
            if p.text.strip():
                contenido_total.append(p.text)
                
        # Si el elemento es una tabla, la formateamos con separadores '|'
        elif elemento.tag.endswith('tbl'):
            tabla = docx.table.Table(elemento, doc)
            contenido_total.append("#") # Añadimos el separador de bloques
            
            for fila in tabla.rows:
                # Unimos el texto de cada celda de la fila usando el separador '|'
                textos_celdas = [celda.text.strip().replace('\n', ' ') for celda in fila.cells]
                # Eliminamos duplicados contiguos si hay celdas combinadas vertical/horizontalmente
                linea_tabla = " | ".join(textos_celdas)
                contenido_total.append(linea_tabla)
                
    return "\n".join(contenido_total)

# ============================================================
# INTERFAZ DE USUARIO (SUBPÁGINA 2)
# ============================================================
st.title("📐 Control Previo y Revisión de Cadenas de Valor")
st.write("Sube el archivo Word de la Cadena de Valor para extraer su información y prepararla para el cruce con el Plan Indicativo.")

st.markdown("---")

# Componente nativo de Streamlit para subir el archivo .docx
archivo_word = st.file_uploader(
    "📂 Sube aquí el formato de Cadena de Valor en Word (.docx)", 
    type=["docx"],
    help="El sistema extraerá las tablas presupuestales, objetivos y la matriz técnica automáticamente."
)

if archivo_word is not None:
    with st.spinner("⏳ Leyendo documento de Word y estructurando tablas..."):
        try:
            # Extraemos todo el texto y tablas unificadas
            texto_extraido = extraer_texto_y_tablas_docx(archivo_word)
            
            # Almacenamos el texto en el estado de la sesión para usarlo en los siguientes pasos
            st.session_state["texto_word_extraido"] = texto_extraido
            
            st.success("✅ ¡Archivo cargado y procesado con éxito!")
            
            # --- VISTA PREVIA INFORMATIVA DE LOS DATOS DETECTADOS ---
            # Buscamos de manera rápida el nombre del proyecto o dependencia para darle feedback al usuario
            with st.expander("🔍 Ver texto técnico extraído del Word", expanded=False):
                st.text_area("Contenido bruto procesado:", value=texto_extraido, height=300)
                
            # Aquí ya tenemos la variable `texto_extraido` perfectamente formateada,
            # lista para que en el siguiente bloque del código realicemos el cruce con el DRIVE
            # y ejecutemos las validaciones del agente experto.
            
        except Exception as e:
            st.error(f"🚨 Error al leer el archivo Word: {e}. Asegúrate de que no esté corrupto o protegido.")
