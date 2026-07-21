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
# COMPONENTE DE AUDITORÍA: WORD ("Indicador de Producto CV - MGA") VS. PI
# ============================================================

st.markdown("---")
st.subheader("🔍 Auditoría de Coherencia: Word vs. Plan Indicativo (PI)")

URL_DRIVE_EXCEL = "https://docs.google.com/spreadsheets/d/18z_tAg7RPvSTSRSTYoYtKgIV3ch3cQ-JbcAOTjQD8ss/export?format=xlsx"

if "df_indicadores_estandar" in st.session_state and not st.session_state["df_indicadores_estandar"].empty:
    df_word = st.session_state["df_indicadores_estandar"].copy()
    
    if st.button("🚀 Ejecutar Cruce y Comparación con PI"):
        with st.spinner("⏳ Conectando con Plan Indicativo y comparando indicadores de producto..."):
            try:
                # 1. Leer el Plan Indicativo en la pestaña "MP" (Encabezados en fila 2)
                df_drive = pd.read_excel(URL_DRIVE_EXCEL, sheet_name="MP", header=1, engine="openpyxl")
                
                columnas_reales = list(df_drive.columns)
                col_codigo_mp = None
                col_indicador_drive = None
                
                # Mapeo de columnas en la pestaña MP del Drive
                for col in columnas_reales:
                    col_limpia = str(col).strip().lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
                    if "codigo mp" in col_limpia:
                        col_codigo_mp = col
                    elif "indicador" in col_limpia and "producto" in col_limpia:
                        col_indicador_drive = col
                
                if col_codigo_mp and col_indicador_drive:
                    
                    # Función para extraer únicamente la secuencia numérica (los dígitos antes del guion o el código)
                    def extraer_codigo_numerico(texto):
                        if pd.isna(texto) or str(texto).strip().lower() == "nan":
                            return ""
                        
                        texto_str = str(texto).strip()
                        
                        # Si tiene guion (-), aislamos la primera parte
                        if "-" in texto_str:
                            texto_str = texto_str.split('-')[0].strip()
                        
                        # Extraer solo los dígitos numéricos
                        digitos = re.sub(r'\D', '', texto_str)
                        return digitos

                    # Indexar el Plan Indicativo por la LLAVE MP
                    dict_pi_indicadores = {}
                    dict_pi_codigos_prod = {}
                    
                    for _, fila in df_drive.iterrows():
                        mp_raw = fila[col_codigo_mp]
                        ind_pi_raw = fila[col_indicador_drive]
                        
                        mp_llave = extraer_codigo_numerico(mp_raw)
                        cod_prod_pi = extraer_codigo_numerico(ind_pi_raw)
                        
                        if mp_llave:
                            dict_pi_indicadores[mp_llave] = str(ind_pi_raw).strip()
                            dict_pi_codigos_prod[mp_llave] = cod_prod_pi

                    # Control y trazabilidad de resultados
                    logs_diagnostico = []
                    resultados_validacion = []
                    textos_pi_encontrados = []
                    codigos_word_extraidos = []
                    codigos_pi_extraidos = []
                    
                    # 2. Iterar sobre las filas extraídas del Word
                    for i, fila_word in df_word.iterrows():
                        mp_word_raw = fila_word.get("Código MP", "")
                        mp_llave_word = extraer_codigo_numerico(mp_word_raw)
                        
                        # CAMBIO SOLICITADO: Lectura desde "Indicador de Producto CV - MGA"
                        ind_word_raw = fila_word.get("Indicador de Producto CV - MGA", "")
                        cod_prod_word = extraer_codigo_numerico(ind_word_raw)
                        
                        codigos_word_extraidos.append(cod_prod_word if cod_prod_word else "No detectado")
                        
                        if not mp_llave_word:
                            resultados_validacion.append("🔴 MP ausente en Word")
                            textos_pi_encontrados.append("N/A")
                            codigos_pi_extraidos.append("N/A")
                            logs_diagnostico.append(f"Fila {i}: Falta el 'Código MP' en el Word.")
                        else:
                            # 3. BÚSQUEDA EN PI CON LA LLAVE MP
                            if mp_llave_word in dict_pi_codigos_prod:
                                cod_prod_pi = dict_pi_codigos_prod[mp_llave_word]
                                texto_pi = dict_pi_indicadores[mp_llave_word]
                                
                                codigos_pi_extraidos.append(cod_prod_pi if cod_prod_pi else "No detectado")
                                textos_pi_encontrados.append(texto_pi)
                                
                                # 4. COMPARACIÓN DE CÓDIGOS EXTRAÍDOS
                                if cod_prod_word and cod_prod_pi and cod_prod_word == cod_prod_pi:
                                    resultados_validacion.append("🟢 Corresponde al PI")
                                    logs_diagnostico.append(
                                        f"✅ Fila {i} (Llave MP: {mp_llave_word}): Coincidencia exacta. "
                                        f"Word: '{cod_prod_word}' == PI: '{cod_prod_pi}'."
                                    )
                                else:
                                    resultados_validacion.append("🔴 Código no coincide")
                                    logs_diagnostico.append(
                                        f"❌ Fila {i} (Llave MP: {mp_llave_word}): DISCREPANCIA DETECTADA. "
                                        f"Word registra código '{cod_prod_word}' en 'Indicador de Producto CV - MGA', "
                                        f"pero el PI tiene asignado el código '{cod_prod_pi}'."
                                    )
                            else:
                                resultados_validacion.append("🔴 MP no existe en PI")
                                textos_pi_encontrados.append("NO EXISTE META EN PI")
                                codigos_pi_extraidos.append("N/A")
                                logs_diagnostico.append(
                                    f"❌ Fila {i}: La llave MP '{mp_llave_word}' (Word) no existe en la pestaña MP del PI."
                                )
                    
                    # Inyección al DataFrame
                    df_word["Cod Indicador Word"] = codigos_word_extraidos
                    df_word["Cod Indicador PI"] = codigos_pi_extraidos
                    df_word["Indicador en PI"] = textos_pi_encontrados
                    df_word["Resultado Validación"] = resultados_validacion
                    
                    # Consola de diagnóstico
                    st.warning("🛠️ **Consola de Diagnóstico: Comparación Indicador de Producto (Word) vs. PI**")
                    with st.expander("👁️ Ver trazabilidad del cruce con el Plan Indicativo", expanded=True):
                        st.code("\n".join(logs_diagnostico), language="text")
                    
                    # Formato semáforo
                    def color_semaforo(val):
                        if "🟢" in str(val):
                            return "background-color: #d4edda; color: #155724; font-weight: bold;"
                        else:
                            return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
                    
                    df_final_render = df_word[[
                        "Código MP", 
                        "Cod Indicador Word",
                        "Indicador de Producto CV - MGA", 
                        "Cod Indicador PI",
                        "Indicador en PI", 
                        "Resultado Validación"
                    ]].copy()
                    
                    st.markdown("##### 📈 Reporte de Inconsistencias: Indicador de Producto Word vs. Plan Indicativo")
                    st.dataframe(
                        df_final_render.style.map(color_semaforo, subset=["Resultado Validación"]),
                        use_container_width=True
                    )
                    
                    conteo_alertas = df_word["Resultado Validación"].str.contains("🔴").sum()
                    if conteo_alertas > 0:
                        st.error(f"⚠️ Se detectaron {conteo_alertas} alertas de inconsistencia entre el Word y el Plan Indicativo.")
                    else:
                        st.success("🎉 ¡Validación correcta! Todos los indicadores corresponden al Plan Indicativo para cada MP.")
                        
                else:
                    st.error("🚨 No se encontraron las columnas necesarias ('Código MP' e 'Indicador de Producto') en la pestaña MP del Plan Indicativo.")
                    
            except Exception as e:
                st.error(f"❌ Error al ejecutar la validación con el Plan Indicativo: {e}")
else:
    st.info("💡 Primero carga el archivo Word para habilitar el cruce contra el Plan Indicativo.")


import streamlit as st
import pandas as pd
import re

import re
import pandas as pd
import streamlit as st

# ============================================================
# MÓDULO DE VERIFICACIÓN Z023: PROYECTOS VIGENCIA 2026
# ============================================================

st.markdown("---")
st.subheader("📦 Auditoría Z023: Proyectos Asociados a la Meta de Producto (Vigencia 2026)")

# Permitir archivos .xlsm (habilitados para macros), .xlsx y .xls
uploaded_z023 = st.file_uploader(
    "Carga el archivo Z023 Consolidado (.xlsm o .xlsx)", 
    type=["xlsm", "xlsx", "xls"], 
    key="z023_uploader"
)

if "df_indicadores_estandar" in st.session_state and not st.session_state["df_indicadores_estandar"].empty:
    df_word = st.session_state["df_indicadores_estandar"].copy()
    
    if uploaded_z023 is not None:
        if st.button("🚀 Analizar Proyectos Aportantes (Z023 - 2026)"):
            with st.spinner("⏳ Procesando Hoja1 de Z023, filtrando vigencia 2026 y agrupando proyectos..."):
                try:
                    # 1. Cargar específicamente la "Hoja1" requerida
                    df_z023 = pd.read_excel(uploaded_z023, sheet_name="Hoja1", engine="openpyxl")
                    
                    # Limpieza flexible de nombres de columnas
                    columnas_z023 = {str(c).strip(): c for c in df_z023.columns}
                    
                    # Función auxiliar para encontrar columnas clave
                    def buscar_columna(patron, diccionario_cols):
                        for c_limpia, c_real in diccionario_cols.items():
                            if patron.lower() in c_limpia.lower():
                                return c_real
                        return None

                    col_vigencia = buscar_columna("Vigencia", columnas_z023)
                    col_cod_mp = buscar_columna("Codigo Meta Producto", columnas_z023) or buscar_columna("Meta Producto", columnas_z023)
                    col_ppm_proj = buscar_columna("PPM: Proyecto", columnas_z023) or buscar_columna("Proyecto", columnas_z023)
                    col_desc_proj = buscar_columna("Descripción PROYECTO", columnas_z023) or buscar_columna("Descripcion", columnas_z023)
                    col_bpin = buscar_columna("Cod.BPIN DNP", columnas_z023) or buscar_columna("BPIN", columnas_z023)

                    if not (col_vigencia and col_cod_mp and col_ppm_proj and col_desc_proj):
                        st.error("🚨 No se encontraron todas las columnas requeridas en la 'Hoja1' del Z023.")
                        st.info(f"Columnas detectadas en Hoja1: {list(df_z023.columns)}")
                    else:
                        # 2. Filtrar únicamente por la Vigencia 2026
                        df_z023_2026 = df_z023[df_z023[col_vigencia].astype(str).str.contains("2026", na=False)].copy()
                        
                        if df_z023_2026.empty:
                            st.warning("⚠️ El archivo Z023 no contiene registros para la vigencia 2026 en la 'Hoja1'.")
                        else:
                            # Helper para aislar los dígitos de la MP concatenados con 'V'
                            def extraer_codigo_numerico(texto):
                                if pd.isna(texto) or str(texto).strip().lower() == "nan":
                                    return ""
                                texto_str = str(texto).strip()
                                if "-" in texto_str:
                                    texto_str = texto_str.split('-')[0].strip()
                                num_extraido = re.sub(r'\D', '', texto_str)
                                return f"V{num_extraido}" if num_extraido else ""

                            # Agrupar Z023 por la MP para consolidar sus proyectos únicos
                            dict_z023_proyectos = {}
                            
                            for _, fila in df_z023_2026.iterrows():
                                mp_key = extraer_codigo_numerico(fila[col_cod_mp])
                                if not mp_key:
                                    continue
                                
                                ppm = str(fila[col_ppm_proj]).strip() if pd.notna(fila[col_ppm_proj]) else "S/C"
                                desc = str(fila[col_desc_proj]).strip() if pd.notna(fila[col_desc_proj]) else "Sin Descripción"
                                bpin = str(fila[col_bpin]).strip() if (col_bpin and pd.notna(fila[col_bpin])) else "No registra"

                                if mp_key not in dict_z023_proyectos:
                                    dict_z023_proyectos[mp_key] = {}

                                # Usamos el código PPM como clave única
                                dict_z023_proyectos[mp_key][ppm] = {
                                    "PPM": ppm,
                                    "Descripcion": desc,
                                    "BPIN": bpin
                                }

                            # 3. Cruzar con las Metas de Producto del Word
                            conteo_proyectos = []
                            listado_codigos_ppm = []
                            resumen_detallado_proyectos = []
                            
                            for i, fila_word in df_word.iterrows():
                                mp_word_raw = fila_word.get("Código MP", "")
                                mp_key_word = extraer_codigo_numerico(mp_word_raw)
                                
                                if mp_key_word in dict_z023_proyectos:
                                    proyectos_map = dict_z023_proyectos[mp_key_word]
                                    cant = len(proyectos_map)
                                    conteo_proyectos.append(cant)
                                    
                                    codigos_ppm = list(proyectos_map.keys())
                                    listado_codigos_ppm.append(", ".join(codigos_ppm))
                                    
                                    detalles = []
                                    for p in proyectos_map.values():
                                        bpin_str = f" (BPIN: {p['BPIN']})" if p['BPIN'] not in ["No registra", "nan", ""] else ""
                                        detalles.append(f"• [{p['PPM']}] {p['Descripcion']}{bpin_str}")
                                    
                                    resumen_detallado_proyectos.append("\n".join(detalles))
                                else:
                                    conteo_proyectos.append(0)
                                    listado_codigos_ppm.append("Sin proyectos 2026")
                                    resumen_detallado_proyectos.append("🔴 No se registraron proyectos asociados en la vigencia 2026 (Z023).")

                            # Inyectar resultados al DataFrame
                            df_word["Cant. Proyectos (2026)"] = conteo_proyectos
                            df_word["Códigos PPM"] = listado_codigos_ppm
                            df_word["Detalle Proyectos Z023"] = resumen_detallado_proyectos
                            
                            # Guardar resultado en session_state para mantener persistencia
                            st.session_state["df_z023_resultado"] = df_word.copy()

                except Exception as e:
                    st.error(f"❌ Error al procesar la 'Hoja1' del archivo Z023: {e}")
    else:
        st.info("📌 Carga el archivo **Z023 Consolidado** para ejecutar el análisis de proyectos 2026.")
    
    # RENDERIZADO PERSISTENTE: Se dibuja si existe en session_state, evitando borrados en reruns
    if "df_z023_resultado" in st.session_state and not st.session_state["df_z023_resultado"].empty:
        st.markdown("##### 📊 Consolidado de Proyectos Aportantes a Metas de Producto (2026)")
        
        df_resumen_render = st.session_state["df_z023_resultado"][[
            "Código MP", 
            "Cant. Proyectos (2026)", 
            "Códigos PPM", 
            "Detalle Proyectos Z023"
        ]].copy()
        
        st.dataframe(df_resumen_render, use_container_width=True)
        
        sin_proyectos = (st.session_state["df_z023_resultado"]["Cant. Proyectos (2026)"] == 0).sum()
        if sin_proyectos > 0:
            st.warning(f"⚠️ Hay {sin_proyectos} Metas de Producto del Word que NO presentan proyectos en la vigencia 2026 del Z023.")
        else:
            st.success("🎉 Todas las Metas de Producto del Word cuentan con al menos un proyecto asignado en la vigencia 2026.")

else:
    st.info("💡 Carga el archivo Word en la sección principal para habilitar el cruce con Z023.")
