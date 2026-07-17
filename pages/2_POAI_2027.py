import streamlit as st
import re
import pandas as pd

def extraer_encabezado_estandar(texto_bruto):
    """
    Usa expresiones regulares para capturar de manera estricta la estructura
    del encabezado institucional definido por el usuario.
    """
    metadatos = {
        "dependencia": None, "fecha": None, "nombre_proyecto": None,
        "id_mga": None, "bpin": None, "codigo_pi": None
    }
    
    # 1. Capturar Dependencia y Fecha
    match_dep = re.search(r"Dependencia Solicitante\s*-\s*Fecha:\s*([\d/]+)", texto_bruto, re.IGNORECASE)
    if match_dep:
        metadatos["fecha"] = match_dep.group(1)
    
    # Buscar el texto que acompaña a la dependencia antes del guion si existe en las primeras líneas
    lineas_inicio = texto_bruto.split("\n")[:10]
    texto_inicio = " ".join(lineas_inicio)
    
    # 2. Capturar Nombre del Proyecto entre comillas tipográficas o normales
    match_nom = re.search(r"PROYECTO INVERSIÓN:\s*[“\"']([^”\"']+)[”\"']", texto_bruto, re.IGNORECASE)
    if match_nom:
        metadatos["nombre_proyecto"] = match_nom.group(1).strip()
        
    # 3. Capturar ID-MGA, BPIN y PIxx-xxxxxx
    match_mga = re.search(r"ID-MGA:\s*(\s*\w+)", texto_bruto, re.IGNORECASE)
    match_bpin = re.search(r"BPIN\s*(\d+)", texto_bruto, re.IGNORECASE)
    match_pi = re.search(r"(PI\d+-\d+)", texto_bruto, re.IGNORECASE)
    
    if match_mga: metadatos["id_mga"] = match_mga.group(1).strip()
    if match_bpin: metadatos["bpin"] = match_bpin.group(1).strip()
    if match_pi: metadatos["codigo_pi"] = match_pi.group(1).strip()
    
    return metadatos

def procesar_tablas_estandar(texto_bruto):
    """
    Divide por bloques '#' para procesar las 5 tablas estándar del formato.
    """
    bloques = [b.strip() for b in texto_bruto.split("#") if b.strip()]
    
    dicc_indicadores = {}
    lista_actividades_poai = []
    
    for bloque in bloques:
        lineas = bloque.split("\n")
        encabezado_tabla = lineas[0].lower()
        
        # --- TABLA 1: DATOS BÁSICOS Y OBJETIVOS ---
        if "no.cv" in encabezado_tabla and "objetivo específico" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|") if p.strip()]
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx not in dicc_indicadores: dicc_indicadores[idx] = {}
                    dicc_indicadores[idx]["dependencia_tabla"] = partes[1] if len(partes) > 1 else ""
                    dicc_indicadores[idx]["nombre_proyecto_tabla"] = partes[2] if len(partes) > 2 else ""
                    dicc_indicadores[idx]["obj_general"] = partes[4] if len(partes) > 4 else ""
                    dicc_indicadores[idx]["obj_especifico"] = partes[5] if len(partes) > 5 else ""

        # --- TABLA 2: ALINEACIÓN PDD ---
        elif "sector mga-sap" in encabezado_tabla and "subprograma plan" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|") if p.strip()]
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        dicc_indicadores[idx]["programa_pdd"] = partes[2] if len(partes) > 2 else ""
                        dicc_indicadores[idx]["meta_resultado"] = partes[4] if len(partes) > 4 else ""

        # --- TABLA 3: PROGRAMACIÓN PLURIANUAL Y CRONOGRAMAS ---
        elif "meta producto plan" in encabezado_tabla and "2027 mga" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|") if p.strip()]
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        # Extraer código limpio MP de la celda
                        celda_mp = partes[1] if len(partes) > 1 else ""
                        match_mp = re.search(r"(MP\d+)", celda_mp)
                        dicc_indicadores[idx]["codigo_mp"] = match_mp.group(1) if match_mp else "SIn Código"
                        dicc_indicadores[idx]["descripcion_mp"] = celda_mp
                        
                        # Metas físicas Plan Indicativo vs MGA
                        dicc_indicadores[idx]["2026_pi"] = partes[4] if len(partes) > 4 else "0"
                        dicc_indicadores[idx]["2027_pi"] = partes[5] if len(partes) > 5 else "0"
                        dicc_indicadores[idx]["2026_mga"] = partes[12] if len(partes) > 12 else "0"
                        dicc_indicadores[idx]["2027_mga"] = partes[13] if len(partes) > 13 else "0"

        # --- TABLA 4: TIPO DE PRODUCTO MGA ---
        elif "observación por indicador" in encabezado_tabla and "tipo prod" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|") if p.strip()]
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        dicc_indicadores[idx]["producto_mga"] = partes[1] if len(partes) > 1 else ""
                        dicc_indicadores[idx]["tipo_producto"] = partes[3] if len(partes) > 3 else "DIRECTO"

        # --- TABLA 5: ACTIVIDADES Y RECURSO POAI 2027 ---
        elif "cod. meta de producto" in encabezado_tabla or "actividad del proyecto" in encabezado_tabla:
            # Buscamos las posiciones de las columnas básicas por si cambian de orden
            columnas = [c.strip().lower() for c in lineas[0].split("|")]
            
            pos_mp = next((i for i, c in enumerate(columnas) if "cod. meta" in c or "producto" in c), 0)
            pos_act = next((i for i, c in enumerate(columnas) if "actividad" in c), 2)
            pos_rec = next((i for i, c in enumerate(columnas) if "recurso" in c or "total 2027" in c), -1)
            
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|") if p.strip()]
                # Validar que sea una fila de contenido presupuestal y no firmas
                if len(partes) >= 3 and "TOTAL" not in partes[0] and "Firma" not in partes[0]:
                    match_mp_act = re.search(r"(MP\d+)", partes[pos_mp])
                    
                    actividad = {
                        "codigo_mp": match_mp_act.group(1) if match_mp_act else "No mapeado",
                        "actividad_descripcion": partes[pos_act] if len(partes) > pos_act else "",
                        "recurso_poai_2027": partes[pos_rec] if pos_rec < len(partes) else "$0"
                    }
                    lista_actividades_poai.append(actividad)

    # Consolidar indicadores cruzados por ID en un DataFrame
    df_indicadores = pd.DataFrame.from_dict(dicc_indicadores, orient="index")
    df_poai = pd.DataFrame(lista_actividades_poai)
    
    return df_indicadores, df_poai

# ============================================================
# RENDERIZADO EN STREAMLIT
# ============================================================
if "texto_word_extraido" in st.session_state:
    texto = st.session_state["texto_word_extraido"]
    
    # Ejecutar extracciones estructuradas
    metadatos = extraer_encabezado_estandar(texto)
    df_ind, df_poai = procesar_tablas_estandar(texto)
    
    st.success("🎯 Datos extraídos bajo la estructura estándar solicitada")
    
    # Panel de control superior (Metadatos del Encabezado)
    with st.expander("📌 Ver Datos de Identificación del Proyecto", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.text_input("Proyecto de Inversión:", value=metadatos["nombre_proyecto"] or "No detectado", disabled=True)
        c2.text_input("Código de Proyecto (PS-SAP / PI):", value=metadatos["codigo_pi"] or "No detectado", disabled=True)
        c3.text_input("Código BPIN:", value=metadatos["bpin"] or "No detectado", disabled=True)
    
    # Mostrar la Matriz Técnica Unificada (Tablas 1 a 4)
    st.markdown("#### 📊 Matriz de Objetivos y Metas Plurianuales (Cruces 1-4)")
    st.dataframe(df_ind, use_container_width=True)
    
    # Mostrar la Matriz de Distribución Presupuestal (Tabla 5)
    st.markdown("#### 💰 Distribución Presupuestal y Actividades POAI 2027 (Tabla 5)")
    st.dataframe(df_poai, use_container_width=True)
    
    # Guardar en estado de sesión para la fase de cruce con el Plan Indicativo
    st.session_state["df_indicadores_estandar"] = df_ind
    st.session_state["df_poai_estandar"] = df_poai
