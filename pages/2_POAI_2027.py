def procesar_tablas_estandar(texto_bruto):
    """
    Segmenta los bloques por '#' y mapea las 5 tablas estándar institucionales
    limpiando de forma estricta los desfaces producidos por los delimitadores '|'.
    """
    bloques = [b.strip() for b in texto_bruto.split("#") if b.strip()]
    
    dicc_indicadores = {}
    lista_actividades_poai = []
    
    for bloque in bloques:
        lineas = bloque.split("\n")
        if not lineas: continue
        encabezado_tabla = lineas[0].lower()
        
        # --- TABLA 1: DATOS BÁSICOS Y OBJETIVOS ---
        # Columnas reales: [0]No.CV | [1]Dependencia | [2]Nombre Proyecto | [3]Fecha CV | [4]Objetivo General Proyecto | [5]Objetivo Específico
        if "no.cv" in encabezado_tabla and "objetivo específico" in encabezado_tabla:
            for l in lineas[1:]:
                # Separamos por '|' y eliminamos espacios extremos
                partes = [p.strip() for p in l.split("|")]
                # Si el conversor dejó un espacio vacío al inicio debido al primer '|', lo removemos
                if partes and partes[0] == "": partes.pop(0)
                
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx not in dicc_indicadores: dicc_indicadores[idx] = {}
                    
                    dicc_indicadores[idx]["Dependencia"] = partes[1] if len(partes) > 1 else ""
                    dicc_indicadores[idx]["Nombre Proyecto"] = partes[2] if len(partes) > 2 else ""
                    dicc_indicadores[idx]["Fecha CV"] = partes[3] if len(partes) > 3 else ""
                    dicc_indicadores[idx]["Objetivo General Proyecto"] = partes[4] if len(partes) > 4 else ""
                    dicc_indicadores[idx]["Objetivo Específico"] = partes[5] if len(partes) > 5 else ""

        # --- TABLA 2: ALINEACIÓN PDD ---
        # Columnas reales: [0]No.CV | [1]Sector MGA-SAP | [2]Línea Estratégica | [3]Programa Plan de Desarrollo | [4]Programa MGA | [5]Meta de Resultado | [6]Subprograma Plan
        elif "sector mga-sap" in encabezado_tabla and "subprograma plan" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|")]
                if partes and partes[0] == "": partes.pop(0)
                
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
        # Columnas: [0]No.CV | [1]Meta Producto Plan | [2]P.G. PI | [3]2024 PI | [4]2025 PI | [5]2026 PI | [6]2027 PI | [7]Código y Nombre Producto Catalogo - MP | ... | [10]2024 MGA | [11]2025 MGA | [12]2026 MGA | [13]2027 MGA
        elif "meta producto plan" in encabezado_tabla and "2027 mga" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|")]
                if partes and partes[0] == "": partes.pop(0)
                
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        celda_mp = partes[1] if len(partes) > 1 else ""
                        match_mp = re.search(r"(MP\d+)", celda_mp)
                        
                        dicc_indicadores[idx]["Código MP"] = match_mp.group(1) if match_mp else "Sin Código"
                        dicc_indicadores[idx]["Descripción Meta Producto"] = celda_mp
                        dicc_indicadores[idx]["2026 PI"] = partes[5] if len(partes) > 5 else "0"
                        dicc_indicadores[idx]["2027 PI"] = partes[6] if len(partes) > 6 else "0"
                        dicc_indicadores[idx]["2026 MGA"] = partes[12] if len(partes) > 12 else "0"
                        dicc_indicadores[idx]["2027 MGA"] = partes[13] if len(partes) > 13 else "0"

        # --- TABLA 4: TIPO DE PRODUCTO MGA ---
        # Columnas: [0]No.CV | [1]Observación por Indicador | [2]Producto CV - MGA | [3]Indicador de Producto CV - MGA | [4]Tipo prod. | [5]Tipo prod2
        elif "observación por indicador" in encabezado_tabla and "tipo prod" in encabezado_tabla:
            for l in lineas[1:]:
                partes = [p.strip() for p in l.split("|")]
                if partes and partes[0] == "": partes.pop(0)
                
                if partes and partes[0].isdigit():
                    idx = int(partes[0])
                    if idx in dicc_indicadores:
                        # Evaluamos ambas celdas de tipo de producto para consolidar en una columna limpia
                        t1 = partes[4] if len(partes) > 4 else ""
                        t2 = partes[5] if len(partes) > 5 else ""
                        dicc_indicadores[idx]["Tipo Producto"] = "INDIRECTO" if "INDIRECTO" in (t1.upper() + t2.upper()) else "DIRECTO"

    df_indicadores = pd.DataFrame.from_dict(dicc_indicadores, orient="index")
    return df_indicadores
