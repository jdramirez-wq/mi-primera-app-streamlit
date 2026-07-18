import streamlit as st
import docx  # Requiere python-docx en requirements.txt
import re
import pandas as pd

st.set_page_config(
    page_title="Revisión de Cadenas de Valor",
    page_icon="📐",
    layout="wide"
)

# ============================================================
# FUNCIONES EXTRACTORAS Y PARSERS CORREGIDOS
# ============================================================

def extraer_texto_y_tablas_docx(file_buffer):
    """Lee el archivo .docx y lo convierte en texto plano estructurado."""
    doc = docx.Document(file_buffer)
    contenido_total = []
    
    for elemento in doc.element.body:
        if elemento.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(elemento, doc)
            if p.text.strip():
                contenido_total.append(p.text)
        elif elemento.tag.endswith('tbl'):
            tabla = docx.table.Table(elemento, doc)
            contenido_total.append("#") 
            for fila in tabla.rows:
                # Extraer texto de cada celda limpiando saltos de línea internos
                textos_celdas = [celda.text.strip().replace('\n', ' ') for celda in fila.cells]
                linea_tabla = " | ".join(textos_celdas)
                contenido_total.append(linea_tabla)
                
    return "\n".join(contenido_total)

def extraer_encabezado_estandar(texto_bruto):
    """Extrae las variables de identificación usando expresiones regulares."""
    metadatos = {
        "dependencia": "No detectada", "fecha": "No detectada", 
        "nombre_proyecto": "No detectado", "id_mga": "No detectado", 
        "bpin": "No detectado", "codigo_pi": "No detectado"
    }
    
    # Capturar Fecha
    match_fecha = re.search(r"Fecha:\s*([\d/:-]+)", texto_bruto, re.IGNORECASE)
    if match_fecha: metadatos["fecha"] = match_fecha.group(1)
    
    # Capturar Nombre del Proyecto
    match_nom = re.search(r"PROYECTO INVERSIÓN:\s*[“\"']([^”\"']+)[”\"']", texto_bruto, re.IGNORECASE)
    if match_nom: metadatos["nombre_proyecto"] = match_nom.group(1).strip()
        
    # Capturar ID-MGA, BPIN y Código PI (PS-SAP)
    match_mga = re.search(r"ID-MGA:\s*(\w+)", texto_bruto, re.IGNORECASE)
    match_bpin = re.search(r"BPIN\s*(\d+)", texto_bruto, re.IGNORECASE)
    match_pi = re.search(r"(PI\d+-\d+)", texto_bruto, re.IGNORECASE)
    
    if match_mga: metadatos["id_mga"] = match_mga.group(1).strip()
    if match_bpin: metadatos["bpin"] = match_bpin.group(1).strip()
    if match_pi: metadatos["codigo_pi"] = match_pi.group(1).strip()
    
    return metadatos

def procesar_tablas_estandar(texto_bruto):
    """Segmenta los bloques por '#' y mapea las 5 tablas de forma tolerante a fallos."""
    bloques = [b.strip() for b in texto_bruto.split("#") if b.strip()]
    
    dicc_indicadores = {}
    lista_actividades_poai = []
    
    for bloque in bloques:
        lineas = bloque.split("\n")
        if not lineas: continue
        encabezado_tabla = lineas[0].lower()
        
        # Función auxiliar para limpiar la línea y extraer las partes útiles
        def limpiar_linea_tabla(l):
            partes = [p.strip() for p in l.split("|")]
            # Filtramos textos vacíos extremos pero conservamos el contenido real
            partes_validas = [p for p in partes if p != ""]
            return partes_validas
        
        # --- TABLA 1: DATOS BÁSICOS Y OBJETIVOS ---
        if "no.cv" in encabezado_tabla and "objetivo específico" in encabezado_tabla:
            for l in lineas[1:]:
                partes = limpiar_linea_tabla(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx not in dicc_indicadores: dicc_indicadores[idx] = {}
                    
                    # Asignación segura verificando el tamaño de la lista
                    dicc_indicadores[idx]["Dependencia"] = partes[1] if len(partes) > 1 else ""
                    dicc_indicadores[idx]["Nombre Proyecto"] = partes[2] if len(partes) > 2 else ""
                    dicc_indicadores[idx]["Fecha CV"] = partes[3] if len(partes) > 3 else ""
                    dicc_indicadores[idx]["Objetivo General Proyecto"] = partes[4] if len(partes) > 4 else ""
                    dicc_indicadores[idx]["Objetivo Específico"] = partes[5] if len(partes) > 5 else ""

        # --- TABLA 2: ALINEACIÓN PDD ---
        elif "sector mga-sap" in encabezado_tabla and "subprograma plan" in encabezado_tabla:
            for l in lineas[1:]:
                partes = limpiar_linea_tabla(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        dicc_indicadores[idx]["Sector MGA-SAP"] = partes[1] if len(partes) > 1 else ""
                        dicc_indicadores[idx]["Línea Estratégica"] = partes[2] if len(partes) > 2 else ""
                        dicc_indicadores[idx]["Programa Plan de Desarrollo"] = partes[3] if len(partes) > 3 else ""
                        dicc_indicadores[idx]["Programa MGA"] = partes[4] if len(partes) > 4 else ""
                        dicc_indicadores[idx]["Meta de Resultado"] = partes[5] if len(partes) > 5 else ""
                        dicc_indicadores[idx]["Subprograma Plan"] = partes[6] if len(partes) > 6 else ""

        # --- TABLA 3: PROGRAMACIÓN PLURIANUAL ---
        elif "meta producto plan" in encabezado_tabla and "2027 mga" in encabezado_tabla:
            for l in lineas[1:]:
                partes = limpiar_linea_tabla(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        celda_mp = partes[1] if len(partes) > 1 else ""
                        match_mp = re.search(r"(MP\d+)", celda_mp)
                        
                        dicc_indicadores[idx]["Código MP"] = match_mp.group(1) if match_mp else "Sin Código"
                        dicc_indicadores[idx]["Descripción Meta Producto"] = celda_mp
                        
                        # Mapeo posicional relativo de los años presupuestales
                        dicc_indicadores[idx]["2026 PI"] = partes[5] if len(partes) > 5 else "0"
                        dicc_indicadores[idx]["2027 PI"] = partes[6] if len(partes) > 6 else "0"
                        dicc_indicadores[idx]["2026 MGA"] = partes[12] if len(partes) > 12 else "0"
                        dicc_indicadores[idx]["2027 MGA"] = partes[13] if len(partes) > 13 else "0"

        # --- TABLA 4: TIPO DE PRODUCTO MGA ---
        elif "observación por indicador" in encabezado_tabla and "tipo prod" in encabezado_tabla:
            for l in lineas[1:]:
                partes = limpiar_linea_tabla(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        # Consolidar el tipo de producto revisando las últimas celdas de la fila
                        texto_tipo = " ".join(partes[3:]).upper()
                        dicc_indicadores[idx]["Tipo Producto"] = "INDIRECTO" if "INDIRECTO" in texto_tipo else "DIRECTO"

        # --- TABLA 5: ACTIVIDADES Y RECURSO POAI 2027 ---
        elif "cod. meta de producto" in encabezado_tabla or "actividad del proyecto" in encabezado_tabla:
            columnas = [c.strip().lower() for c in lineas[0].split("|") if c.strip()]
            
            pos_mp = next((i for i, c in enumerate(columnas) if "cod. meta" in c or "producto" in c), 0)
            pos_act = next((i for i, c in enumerate(columnas) if "actividad" in c), 1)
            pos_rec = next((i for i, c in enumerate(columnas) if "recurso" in c or "total 2027" in c), -1)
            
            for l in lineas[1:]:
                partes = limpiar_linea_tabla(l)
                if len(partes) >= 2 and "TOTAL" not in partes[0].upper() and "FIRMA" not in partes[0].upper():
                    match_mp_act = re.search(r"(MP\d+)", partes[pos_mp])
                    actividad = {
                        "Código MP": match_mp_act.group(1) if match_mp_act else "No mapeado",
                        "Actividad Descripción": partes[pos_act] if len(partes) > pos_act else "",
                        "Recurso POAI 2027": partes[pos_rec] if pos_rec < len(partes) else "$0"
                    }
                    lista_actividades_poai.append(actividad)

    df_indicadores = pd.DataFrame.from_dict(dicc_indicadores, orient="index")
    df_poai = pd.DataFrame(lista_actividades_poai)
    
    return df_indicadores, df_poai

# ============================================================
# INTERFAZ DE USUARIO DE STREAMLIT (ENTRADA Y RENDER)
# ============================================================

st.title("📐 Control Previo y Revisión de Cadenas de Valor")
st.write("Sube el archivo Word oficial para estructurar las variables y prepararlas para el cruce.")

st.markdown("---")

archivo_word = st.file_uploader(
    "📂 Sube aquí el formato de Cadena de Valor en Word (.docx)", 
    type=["docx"]
)

if archivo_word is not None:
    with st.spinner("⏳ Procesando documento y estructurando tablas estándar..."):
        try:
            texto_extraido = extraer_texto_y_tablas_docx(archivo_word)
            st.session_state["texto_word_extraido"] = texto_extraido
            
            metadatos = extraer_encabezado_estandar(texto_extraido)
            df_ind, df_poai = procesar_tablas_estandar(texto_extraido)
            
            st.success("✅ ¡Archivo procesado con éxito!")
            
            st.markdown("### 📌 Identificación del Proyecto")
            c1, c2, c3 = st.columns(3)
            c1.text_input("Proyecto de Inversión:", value=metadatos["nombre_proyecto"], disabled=True)
            c2.text_input("Código de Proyecto (PS-SAP):", value=metadatos["codigo_pi"], disabled=True)
            c3.text_input("Código BPIN:", value=metadatos["bpin"], disabled=True)
            
            st.markdown("### 📊 Matriz Técnica de Objetivos y Metas Plurianuales")
            if not df_ind.empty:
                st.dataframe(df_ind, use_container_width=True)
            else:
                st.warning("No se pudieron mapear las tablas con la estructura técnica esperada.")
            
            st.markdown("### 💰 Distribución Presupuestal y Actividades POAI 2027")
            if not df_poai.empty:
                st.dataframe(df_poai, use_container_width=True)
            else:
                st.warning("No se detectaron actividades financieras en la sección presupuestal.")
                
        except Exception as e:
            st.error(f"🚨 Error crítico en el parseo del archivo: {e}")
