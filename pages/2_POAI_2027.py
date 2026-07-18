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
# FUNCIONES EXTRACTORAS Y PARSERS INSTITUCIONALES
# ============================================================

def extraer_texto_y_tablas_docx(file_buffer):
    """Lee el archivo .docx y lo convierte en texto plano estructurado por bloques."""
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
                textos_celdas = [celda.text.strip().replace('\n', ' ') for celda in fila.cells]
                linea_tabla = " | ".join(textos_celdas)
                contenido_total.append(linea_tabla)
                
    return "\n".join(contenido_total)

def extraer_encabezado_estandar(texto_bruto):
    """Extrae las variables de identificación del encabezado."""
    metadatos = {
        "dependencia": "No detectada", "fecha": "No detectada", 
        "nombre_proyecto_encabezado": "No detectado", "id_mga": "No detectado", 
        "bpin": "No detectado", "codigo_pi": "No detectado"
    }
    match_fecha = re.search(r"Fecha:\s*([\d/:-]+)", texto_bruto, re.IGNORECASE)
    if match_fecha: metadatos["fecha"] = match_fecha.group(1)
    
    match_nom = re.search(r"PROYECTO INVERSIÓN:\s*[“\"']([^”\"']+)[”\"']", texto_bruto, re.IGNORECASE)
    if match_nom: metadatos["nombre_proyecto_encabezado"] = match_nom.group(1).strip()
        
    match_mga = re.search(r"ID-MGA:\s*(\w+)", texto_bruto, re.IGNORECASE)
    match_bpin = re.search(r"BPIN\s*(\d+)", texto_bruto, re.IGNORECASE)
    match_pi = re.search(r"(PI\d+-\d+)", texto_bruto, re.IGNORECASE)
    
    if match_mga: metadatos["id_mga"] = match_mga.group(1).strip()
    if match_bpin: metadatos["bpin"] = match_bpin.group(1).strip()
    if match_pi: metadatos["codigo_pi"] = match_pi.group(1).strip()
    
    return metadatos

def procesar_tablas_estandar(texto_bruto):
    """Segmenta los bloques y procesa las 5 tablas de forma estricta según las columnas indicadas."""
    bloques = [b.strip() for b in texto_bruto.split("#") if b.strip()]
    
    dicc_indicadores = {}
    lista_actividades_poai = []
    recurso_total_proyecto = "$0"
    nombre_proyecto_tabla = "No detectado"
    
    for bloque in bloques:
        lineas = bloque.split("\n")
        if not lineas: continue
        encabezado_tabla = lineas[0].lower()
        
        def descomponer_linea(l):
            return [p.strip() for p in l.split("|")]
        
        # --- TABLA 1: DATOS BÁSICOS Y OBJETIVOS ---
        if "no.cv" in encabezado_tabla and "objetivo general proyecto" in encabezado_tabla:
            for l in lineas[1:]:
                partes = descomponer_linea(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx not in dicc_indicadores: dicc_indicadores[idx] = {}
                    dicc_indicadores[idx]["No.CV"] = partes[1] if len(partes) > 1 else ""
                    dicc_indicadores[idx]["Dependencia"] = partes[2] if len(partes) > 2 else ""
                    dicc_indicadores[idx]["Nombre Proyecto"] = partes[3] if len(partes) > 3 else ""
                    dicc_indicadores[idx]["Fecha CV"] = partes[4] if len(partes) > 4 else ""
                    dicc_indicadores[idx]["Objective General Proyecto"] = partes[5] if len(partes) > 5 else ""
                    dicc_indicadores[idx]["Objetivo Específico"] = partes[6] if len(partes) > 6 else ""
                    
                    # Extraemos el nombre del proyecto de la primera fila válida encontrada
                    if nombre_proyecto_tabla == "No detectado" and partes[3]:
                        nombre_proyecto_tabla = partes[3]

        # --- TABLA 2: ALINEACIÓN PDD ---
        elif "sector mga-sap" in encabezado_tabla and "subprograma plan" in encabezado_tabla:
            for l in lineas[1:]:
                partes = descomponer_linea(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        dicc_indicadores[idx]["Sector MGA-SAP"] = partes[1] if len(partes) > 1 else ""
                        dicc_indicadores[idx]["Línea Estratégica"] = partes[2] if len(partes) > 2 else ""
                        dicc_indicadores[idx]["Programa Plan de Desarrollo"] = partes[3] if len(partes) > 3 else ""
                        dicc_indicadores[idx]["Programa MGA"] = partes[4] if len(partes) > 4 else ""
                        dicc_indicadores[idx]["Meta de Resultado"] = partes[5] if len(partes) > 5 else ""
                        dicc_indicadores[idx]["Subprograma Plan"] = partes[6] if len(partes) > 6 else ""

        # --- TABLA 3: PROGRAMACIÓN PLURIANUAL + EXTRACCIÓN CÓDIGO MP ---
        elif "meta producto plan" in encabezado_tabla and "meta total mga" in encabezado_tabla:
            for l in lineas[1:]:
                partes = descomponer_linea(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        meta_producto_texto = partes[1] if len(partes) > 1 else ""
                        dicc_indicadores[idx]["Meta Producto Plan"] = meta_producto_texto
                        
                        # EXTRACCIÓN ADICIONAL: Aislar el Código MP usando Expresiones Regulares
                        match_mp = re.search(r"(MP\d+)", meta_producto_texto)
                        dicc_indicadores[idx]["Código MP"] = match_mp.group(1) if match_mp else "Sin Código"
                        
                        dicc_indicadores[idx]["P.G. PI"] = partes[2] if len(partes) > 2 else ""
                        dicc_indicadores[idx]["2024 PI"] = partes[3] if len(partes) > 3 else ""
                        dicc_indicadores[idx]["2025 PI"] = partes[4] if len(partes) > 4 else ""
                        dicc_indicadores[idx]["2026 PI"] = partes[5] if len(partes) > 5 else ""
                        dicc_indicadores[idx]["2027 PI"] = partes[6] if len(partes) > 6 else ""
                        dicc_indicadores[idx]["Código y Nombre Producto Catalogo - MP"] = partes[7] if len(partes) > 7 else ""
                        dicc_indicadores[idx]["Indicador de Producto Catalogo - MP"] = partes[8] if len(partes) > 8 else ""
                        dicc_indicadores[idx]["Unidad de Medida"] = partes[9] if len(partes) > 9 else ""
                        dicc_indicadores[idx]["Meta Total MGA"] = partes[10] if len(partes) > 10 else ""
                        dicc_indicadores[idx]["2024 MGA"] = partes[11] if len(partes) > 11 else ""
                        dicc_indicadores[idx]["2025 MGA"] = partes[12] if len(partes) > 12 else ""
                        dicc_indicadores[idx]["2026 MGA"] = partes[13] if len(partes) > 13 else ""
                        dicc_indicadores[idx]["2027 MGA"] = partes[14] if len(partes) > 14 else ""

        # --- TABLA 4: TIPO DE PRODUCTO MGA ---
        elif "observación por indicador mga" in encabezado_tabla and "producto cv - mga" in encabezado_tabla:
            for l in lineas[1:]:
                partes = descomponer_linea(l)
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        dicc_indicadores[idx]["Observación por Indicador MGA - Formulador"] = partes[1] if len(partes) > 1 else ""
                        dicc_indicadores[idx]["Producto CV - MGA"] = partes[2] if len(partes) > 2 else ""
                        dicc_indicadores[idx]["Indicador de Producto CV - MGA"] = partes[3] if len(partes) > 3 else ""
                        dicc_indicadores[idx]["Tipo prod."] = partes[4] if len(partes) > 4 else ""
                        dicc_indicadores[idx]["Tipo prod2"] = partes[5] if len(partes) > 5 else ""

        # --- TABLA 5: POAI 2027 Y FILA DE TOTALES ---
        elif "cod. meta de producto" in encabezado_tabla or "actividad del proyecto" in encabezado_tabla:
            columnas = [c.strip().lower() for c in lineas[0].split("|") if c.strip()]
            pos_mp = next((i for i, c in enumerate(columnas) if "cod. meta" in c or "producto" in c), 0)
            pos_prod = next((i for i, c in enumerate(columnas) if "producto mga" in c), 1)
            pos_act = next((i for i, c in enumerate(columnas) if "actividad" in c), 2)
            pos_rec = next((i for i, c in enumerate(columnas) if "recurso total" in c or "total 2027" in c), -1)
            
            for l in lineas[1:]:
                partes = descomponer_linea(l)
                texto_linea = " ".join(partes).upper()
                
                if "TOTAL RECURSOS 2027" in texto_linea or "PROYECTO DE INVERSIÓN" in texto_linea:
                    recurso_total_proyecto = partes[-1] if partes else "$0"
                    continue
                
                if len(partes) >= 3 and "FIRMA" not in partes[0].upper() and partes[0] != "":
                    match_mp_act = re.search(r"(MP\d+)", partes[pos_mp])
                    actividad = {
                        "COD. META DE PRODUCTO": partes[pos_mp] if len(partes) > pos_mp else "",
                        "Código MP Extrayendo": match_mp_act.group(1) if match_mp_act else "Sin Código",
                        "PRODUCTO MGA (COD+TEXTO)": partes[pos_prod] if len(partes) > pos_prod else "",
                        "ACTIVIDAD DEL PROYECTO": partes[pos_act] if len(partes) > pos_act else "",
                        "RECURSO TOTAL 2027": partes[pos_rec] if pos_rec < len(partes) else "$0"
                    }
                    lista_actividades_poai.append(actividad)

    df_indicadores = pd.DataFrame.from_dict(dicc_indicadores, orient="index")
    
    # Reorganizar columnas para dejar el "Código MP" en una ubicación visible e importante
    if not df_indicadores.empty and "Código MP" in df_indicadores.columns:
        cols = list(df_indicadores.columns)
        cols.insert(0, cols.pop(cols.index("Código MP")))
        df_indicadores = df_indicadores[cols]
        
    df_poai = pd.DataFrame(lista_actividades_poai)
    
    return df_indicadores, df_poai, recurso_total_proyecto, nombre_proyecto_tabla

# ============================================================
# INTERFAZ DE USUARIO DE STREAMLIT
# ============================================================

st.title("📐 Control Previo y Revisión de Cadenas de Valor")
st.write("Carga el archivo Word técnico para procesar de forma estricta los campos del Plan de Desarrollo.")

st.markdown("---")

archivo_word = st.file_uploader("📂 Sube aquí el formato de Cadena de Valor en Word (.docx)", type=["docx"])

if archivo_word is not None:
    with st.spinner("⏳ Mapeando columnas y consolidando matrices..."):
        try:
            texto_extraido = extraer_texto_y_tablas_docx(archivo_word)
            st.session_state["texto_word_extraido"] = texto_extraido
            
            metadatos = extraer_encabezado_estandar(texto_extraido)
            df_ind, df_poai, total_presupuesto, proyecto_nombre_tabla = procesar_tablas_estandar(texto_extraido)
            
            st.success("✅ ¡Archivo procesado con éxito!")
            
            # Bloque 1: Datos de identificación usando el nombre extraído de la columna del proyecto
            st.markdown("### 📌 Identificación del Proyecto")
            c1, c2, c3 = st.columns(3)
            c1.text_input("Proyecto de Inversión (Extraído de la Tabla):", value=proyecto_nombre_tabla, disabled=True)
            c2.text_input("Código de Proyecto (PS-SAP):", value=metadatos["codigo_pi"], disabled=True)
            c3.text_input("Código BPIN:", value=metadatos["bpin"], disabled=True)
            
            # Bloque 2: Matriz Unificada de Indicadores con la nueva columna Código MP posicionada al inicio
            st.markdown("### 📊 Matriz Completa de Indicadores y Objetivos (Tablas 1-4)")
            if not df_ind.empty:
                st.dataframe(df_ind, use_container_width=True)
            else:
                st.warning("No se pudo estructurar la matriz técnica de indicadores.")
            
            # Bloque 3: Plan Operativo Anual de Inversiones (Tabla 5)
            st.markdown("### 💰 Distribución Presupuestal y Actividades POAI 2027")
            if not df_poai.empty:
                st.dataframe(df_poai, use_container_width=True)
                st.metric(label="💰 TOTAL RECURSOS 2027 – PROYECTO DE INVERSIÓN", value=total_presupuesto)
            else:
                st.warning("No se detectaron actividades presupuestales en la Tabla 5.")
                
            # Guardado en sesión listo para cruces analíticos
            st.session_state["df_indicadores_estandar"] = df_ind
            st.session_state["df_poai_estandar"] = df_poai
            st.session_state["total_presupuesto_poai"] = total_presupuesto
            
        except Exception as e:
            st.error(f"🚨 Error en el procesamiento del documento: {e}")

# ============================================================
# COMPONENTE DE AUDITORÍA: VALIDACIÓN DE NUMERACIÓN DE INDICADORES
# ============================================================

st.markdown("---")
st.subheader("🔍 Auditoría de Coherencia: Word vs. Plan Indicativo (Drive)")

URL_DRIVE_EXCEL = "https://docs.google.com/spreadsheets/d/18z_tAg7RPvSTSRSTYoYtKgIV3ch3cQ-JbcAOTjQD8ss/export?format=xlsx"

if "df_indicadores_estandar" in st.session_state and not st.session_state["df_indicadores_estandar"].empty:
    df_word = st.session_state["df_indicadores_estandar"].copy()
    
    if st.button("🚀 Ejecutar Cruce y Comparación de Indicadores"):
        with st.spinner("⏳ Localizando filas por MP y comparando numeración de indicadores..."):
            try:
                # 1. Leer el archivo Excel (Pestaña MP, encabezados en fila 2)
                df_drive = pd.read_excel(URL_DRIVE_EXCEL, sheet_name="MP", header=1, engine="openpyxl")
                
                columnas_reales = list(df_drive.columns)
                col_codigo_mp = None
                col_indicador_drive = None
                
                for col in columnas_reales:
                    col_limpia = str(col).strip().lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
                    if "codigo mp" in col_limpia:
                        col_codigo_mp = col
                    elif "indicador" in col_limpia and "producto" in col_limpia:
                        col_indicador_drive = col
                
                if col_codigo_mp and col_indicador_drive:
                    
                    # Función para aislar el código de la MP por el guion
                    def aislar_codigo_por_guion(valor):
                        if pd.isna(valor) or str(valor).strip().lower() == "nan":
                            return ""
                        partes = str(valor).split('-')
                        return str(partes[0]).strip()

                    # --- NUEVA FUNCIÓN: EXTRAE SOLO LOS NÚMEROS DEL INDICADOR ---
                    # Si el indicador dice "Indicador 2. Porcentaje de...", extrae "2"
                    def extraer_numeracion_indicador(texto):
                        if pd.isna(texto) or str(texto).strip().lower() == "nan":
                            return ""
                        # Busca el primer bloque de números pequeños (1 a 3 dígitos) que suele identificar al número de indicador
                        match = re.search(r'\b\d{1,3}\b', str(texto))
                        return match.group(0).strip() if match else ""

                    # Mapeamos el Drive indexando por el código de la MP limpia
                    dict_drive_indicadores = {}
                    for _, fila in df_drive.iterrows():
                        cod_raw = fila[col_codigo_mp]
                        ind_text = fila[col_indicador_drive]
                        cod_limpio = aislar_codigo_por_guion(cod_raw)
                        
                        if cod_limpio:
                            # Guardamos el texto completo del indicador del Drive indexado por su código MP
                            dict_drive_indicadores[cod_limpio] = str(ind_text).strip()

                    # Listas para almacenar la trazabilidad y resultados
                    logs_diagnostico = []
                    resultados_validacion = []
                    indicadores_drive_encontrados = []
                    numeros_word = []
                    numeros_drive = []
                    
                    # 2. Iterar sobre las filas del Word
                    for i, fila_word in df_word.iterrows():
                        cod_word_raw = fila_word.get("Código MP", "")
                        cod_word_limpio = aislar_codigo_por_guion(cod_word_raw)
                        texto_indicador_word = fila_word.get("Indicador de Producto CV - MGA", "")
                        
                        # Extraer el número del indicador del Word (Ej: "1" o "2")
                        num_ind_word = extraer_numeracion_indicador(texto_indicador_word)
                        numeros_word.append(num_ind_word if num_ind_word else "No detectado")
                        
                        if not cod_word_limpio:
                            resultados_validacion.append("🔴 MP ausente en Word")
                            indicadores_drive_encontrados.append("N/A")
                            numeros_drive.append("N/A")
                            logs_diagnostico.append(f"Fila {i}: Código MP vacío en Word.")
                        else:
                            # Buscar la fila correspondiente en el Drive usando la MP
                            if cod_word_limpio in dict_drive_indicadores:
                                texto_indicador_drive = dict_drive_indicadores[cod_word_limpio]
                                indicadores_drive_encontrados.append(texto_indicador_drive)
                                
                                # Extraer el número del indicador del Drive
                                num_ind_drive = extraer_numeracion_indicador(texto_indicador_drive)
                                numeros_drive.append(num_ind_drive if num_ind_drive else "No detectado")
                                
                                # --- LA VERDADERA COMPARACIÓN BASADA EN NÚMEROS ---
                                if num_ind_word and num_ind_drive and num_ind_word == num_ind_drive:
                                    resultados_validacion.append("🟢 Coincide Numeración")
                                    logs_diagnostico.append(
                                        f"✅ Fila {i} (MP {cod_word_limpio}): Coincidencia exacta. "
                                        f"Número Word: '{num_ind_word}' == Número Drive: '{num_ind_drive}'."
                                    )
                                else:
                                    resultados_validacion.append("🔴 Diferente Indicador")
                                    logs_diagnostico.append(
                                        f"❌ Fila {i} (MP {cod_word_limpio}): ¡FALSO POSITIVO EVITADO! "
                                        f"La MP existe, pero el indicador en Word tiene número '{num_ind_word}' "
                                        f"y en el Drive corresponde al número '{num_ind_drive}'."
                                    )
                            else:
                                resultados_validacion.append("🔴 MP no existe en Drive")
                                indicadores_drive_encontrados.append("NO EXISTE META")
                                numeros_drive.append("N/A")
                                logs_diagnostico.append(f"❌ Fila {i}: La Meta de Producto '{cod_word_limpio}' no se encuentra en el Drive.")
                    
                    # Inyectar las nuevas métricas al DataFrame de visualización
                    df_word["Indicador en Drive"] = indicadores_drive_encontrados
                    df_word["Num Word"] = numeros_word
                    df_word["Num Drive"] = numeros_drive
                    df_word["Resultado Validación"] = resultados_validacion
                    
                    # Desplegar la consola de trazabilidad de inmediato
                    st.warning("🛠️ **Consola de Análisis de Numeración Interna:**")
                    with st.expander("👁️ Ver comparación detallada de números de indicadores", expanded=True):
                        st.code("\n".join(logs_diagnostico), language="text")
                    
                    # Formato semáforo para la tabla
                    def color_semaforo(val):
                        if "🟢" in str(val):
                            return "background-color: #d4edda; color: #155724; font-weight: bold;"
                        else:
                            return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
                    
                    # Proyectar el DataFrame final ordenado para el analista
                    df_final_render = df_word[[
                        "Código MP", 
                        "Num Word",
                        "Indicador de Producto CV - MGA", 
                        "Num Drive",
                        "Indicador en Drive", 
                        "Resultado Validación"
                    ]].copy()
                    
                    st.markdown("##### 📈 Reporte de Consistencia de Indicadores (Filtro por Numeración)")
                    st.dataframe(
                        df_final_render.style.map(color_semaforo, subset=["Resultado Validación"]),
                        use_container_width=True
                    )
                    
                    conteo_errores = df_word["Resultado Validación"].str.contains("🔴").sum()
                    if conteo_errores > 0:
                        st.error(f"⚠️ Se detectaron {conteo_errores} diferencias entre el indicador asignado en Word y el Plan Indicativo del Drive.")
                    else:
                        st.success("🎉 ¡Excelente! Todos los indicadores coinciden plenamente en su numeración con el Drive.")
                        
                else:
                    st.error("🚨 No se pudieron mapear las columnas 'Código MP' e 'Indicador de Producto' en el Drive.")
                    
            except Exception as e:
                st.error(f"❌ Error en el procesamiento del cruce por numeración: {e}")
else:
    st.info("💡 Por favor, primero carga un archivo Word en la sección superior.")
