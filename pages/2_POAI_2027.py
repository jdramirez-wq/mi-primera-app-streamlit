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

import streamlit as st
import re
import pandas as pd

def consolidar_informacion_word(texto_bruto):
    """
    Parsea las líneas del Word extraído, unifica las tablas fragmentadas 
    por medio del ID 'No.CV' y extrae metadatos clave.
    """
    lineas = texto_bruto.split("\n")
    
    # Diccionario intermedio para indexar por el número de fila (1, 2, 3, 4)
    datos_consolidados = {}
    
    # Diccionario para almacenar metadatos generales del proyecto
    metadatos = {
        "dependencia": "No detectada",
        "nombre_proyecto": "No detectado",
        "bpin": "No detectado",
        "codigo_pi": "No detectado"
    }
    
    for linea in lineas:
        linea_limpia = linea.strip()
        if not linea_limpia or linea_limpia.startswith("#") or "No.CV" in linea_limpia:
            continue
            
        # 1. Procesamiento de filas de indicadores (Empiezan con número | )
        match_fila = re.match(r"^(\d+)\s*\|\s*(.*)", linea_limpia)
        if match_fila:
            id_cv = int(match_fila.group(1))
            celdas = [c.strip() for c in match_fila.group(2).split("|")]
            
            if id_cv not in datos_consolidados:
                datos_consolidados[id_cv] = {}
                
            # Identificar dinámicamente qué bloque de columnas estamos procesando según el contenido
            if "SECRETARÍA" in celdas[0] or "Gobernación" in celdas[1]:
                # Bloque 1: Metadatos del proyecto y objetivos
                metadatos["dependencia"] = celdas[0]
                metadatos["nombre_proyecto"] = celdas[1]
                datos_consolidados[id_cv]["objetivo_especifico"] = celdas[4] if len(celdas) > 4 else ""
                
            elif "MGA:" in celdas[0] or "Valle competitivo" in celdas[1]:
                # Bloque 2: Línea, Programa y Meta de Resultado
                datos_consolidados[id_cv]["meta_resultado"] = celdas[3]
                
            elif any("MP14" in c for c in celdas):
                # Bloque 3: Códigos MP y Programación Plurianual MGA
                # Buscamos cuál celda contiene el patrón de la Meta de Producto
                celda_mp = next((c for c in celdas if "MP14" in c), "")
                match_mp = re.search(r"(MP\d+)", celda_mp)
                
                datos_consolidados[id_cv]["codigo_mp"] = match_mp.group(1) if match_mp else "No encontrado"
                datos_consolidados[id_cv]["descripcion_mp"] = celda_mp
                # Extraer las metas físicas anuales del proyecto (MGA) de las últimas columnas
                if len(celdas) >= 13:
                    datos_consolidados[id_cv]["meta_2026_mga"] = celdas[11]
                    datos_consolidados[id_cv]["meta_2027_mga"] = celdas[12]
                    
            elif "DIRECTO" in celdas or "INDIRECTO" in celdas:
                # Bloque 4: Tipo de producto e indicador MGA
                datos_consolidados[id_cv]["producto_mga_cv"] = celdas[0]
                datos_consolidados[id_cv]["tipo_producto"] = "DIRECTO" if "DIRECTO" in celdas else "INDIRECTO"

        # 2. Extracción de códigos BPIN y PI desde el texto de la Observación General
        if "BPIN" in linea_limpia or "PI33" in linea_limpia:
            match_bpin = re.search(r"BPIN\s*(\d+)", linea_limpia)
            match_pi = re.search(r"(PI\d+-\d+)", linea_limpia)
            if match_bpin: metadatos["bpin"] = match_bpin.group(1)
            if match_pi: metadatos["codigo_pi"] = match_pi.group(1)
            
    # Convertir el diccionario unificado a un DataFrame limpio de Pandas
    df_proyecto = pd.DataFrame.from_dict(datos_consolidados, orient="index")
    return metadatos, df_proyecto

# ============================================================
# INTEGRACIÓN EN LA INTERFAZ DE STREAMLIT
# ============================================================
if "texto_word_extraido" in st.session_state:
    texto = st.session_state["texto_word_extraido"]
    
    # Procesar de inmediato el texto para estructurarlo
    metadatos, df_proyecto = consolidar_informacion_word(texto)
    
    st.markdown("---")
    st.subheader("📊 Datos del Proyecto unificados automáticamente")
    
    # Mostrar tarjetas de información general extraída
    col1, col2, col3 = st.columns(3)
    col1.metric("Código del Proyecto", metadatos["codigo_pi"])
    col2.metric("Código BPIN", metadatos["bpin"])
    col3.metric("Indicadores a Evaluar", len(df_proyecto))
    
    st.write(f"**Dependencia Solicitante:** {metadatos['dependencia']}")
    st.write(f"**Nombre del Proyecto:** {metadatos['nombre_proyecto']}")
    
    # Mostrar la tabla consolidated final
    st.markdown("##### Matriz Unificada de Metas de Producto (Cadenas de Valor)")
    st.dataframe(df_proyecto, use_container_width=True)
    
    st.session_state["df_proyecto_word"] = df_proyecto
