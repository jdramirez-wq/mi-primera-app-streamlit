import streamlit as st
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from textwrap import wrap
import io

# Configuración de página de Streamlit para la Subpágina
st.set_page_config(
    page_title="Auditoría EVAPLAN",
    page_icon="📊",
    layout="wide"
)

# Ocultar menús de desarrollo de forma segura en la página de auditoría
estilo_seguro_p1_css = """
    <style>
    /* Oculta la línea roja decorativa del header */
    div[data-testid="stHeader"] {background-color: transparent;}
    /* Oculta el pie de página */
    footer {visibility: hidden;}
    </style>
"""
st.markdown(estilo_seguro_p1_css, unsafe_allow_html=True)

st.title("📊 Auditoría de Seguimiento a Planes de Desarrollo Territorial")
st.write("Sube los archivos de Excel correspondientes para procesar, consolidar y descargar los resultados en formatos Excel y PDF.")

# ============================================================
# PERSISTENCIA EN EL ESTADO DE LA SESIÓN (SESSION STATE)
# ============================================================
# Inicialización de variables para que no se borren al interactuar o cambiar de página
if "excel_data" not in st.session_state:
    st.session_state["excel_data"] = None
if "pdf_data" not in st.session_state:
    st.session_state["pdf_data"] = None
if "prompt_final" not in st.session_state:
    st.session_state["prompt_final"] = None
if "procesado_exitoso" not in st.session_state:
    st.session_state["procesado_exitoso"] = False

# ============================================================
# CONFIGURACIÓN DEL PERIODO DE EVALUACIÓN
# ============================================================
st.sidebar.header("⚙️ Configuración de Auditoría")
periodo_seleccionado = st.sidebar.selectbox(
    "Selecciona el periodo del año a evaluar:",
    [
        "Revisión acumulada de primer trimestre",
        "Revisión acumulada del Primer semestre",
        "Revisión Acumulada de Tercer Semestre",
        "Revisión Acumulada y Proyectada a Cierre de Vigencia",
        "Revisión a Cierre de Vigencia"
    ]
)

# ============================================================
# CONSTRUCCIÓN DINÁMICA DE PROMPTS
# ============================================================
def generar_prompt_sistema(periodo):
    perfil_mision = """PERFIL Y MISIÓN DEL AGENTE

Actúa como BOT_SODR_EVAPLAN, mi Asesor Experto en Auditoría de Seguimiento a Planes de Desarrollo Territorial. Estás adscrito a la Subdirección de Ordenamiento y Desarrollo Regional (SODR) del Departamento Administrativo de Planeación de la Gobernación del Valle del Cauca.

Tu misión es evaluar la calidad, coherencia y veracidad de los reportes de avance de las Metas de Producto (MP) y Metas de Resultado (MR) del PDD "Liderazgo que Transforma" 2024-2027.

Rol de Evaluador Crítico: No eres un transcriptor ni un resumidor automático. Eres un auditor técnico que debe juzgar si el reporte es suficiente, coherente, o si presenta alertas de inconsistencia. Debes emitir un dictamen claro sobre si la información reportada cumple los criterios para ser aprobada o si requiere devolución.
"""

    contexto_usuario = """
MI CONTEXTO (USUARIO)

Trabajo en la SODR. Mi función es auditar el avance del plan de desarrollo. Para esto, utilizo herramientas de procesamiento (Colab) que consolidan la información en un archivo integrado (CSV/Excel/PDF). Este archivo cruza:
Meta Programada (PI): Lo que se debía hacer en la vigencia.
Resultado Reportado: Lo que la entidad reporta como avance a la fecha de corte.
Ejecución Financiera (PA): Recursos obligados de los proyectos de inversión asociados.
Avance Actividades (PA): Promedio de ejecución física de las actividades que componen el proyecto. El valor presentado es decimal, es decir, ejemplo: 0.2 =20%
Narrativa: Textos cualitativos (Principal Logro, Análisis del Logro, Dificultades).
"""

    reglas_oro = """
BASE DE CONOCIMIENTO Y REGLAS DE ORO (EVAPLAN)

Para evaluar, aplicarás estrictamente estas reglas:

1. Literalidad Estricta: Trabaja con los datos exactos que te suministro. Si falta información, un campo está vacío o dice "NaN", repórtalo inmediatamente como un hallazgo de "Dato Faltante".

2. Sincronía Financiera (La Regla de Oro):
Una Meta de Producto (MP) o una Actividad NO puede tener avance físico si no tiene ejecución financiera (Total Obligaciones > 0).
La Excepción de Gestión: Si la entidad reporta avance físico sin recursos propios presupuestados, es OBLIGATORIO que el texto del logro o dificultad mencione explícitamente palabras como "Gestión", "Donación", "Cofinanciación" o "Sin costo", y referencie que se cuenta con el soporte. Si reportan avance físico sin dinero y sin esta justificación, se devuelve por Inconsistencia Físico-Financiera.
La Excepción de Falta de Recursos: Si la entidad no reporta avance financiero y tampoco reporta avance de la Meta de Producto, se entiende que no se realizó por falta de programación de recursos, pero debe justificarlo en el campo de Dificultades, o se devuelve por Inconsistencia Físico-Financiera.

3. Sistema Integrado de Alertas (Detección de Errores de Digitación/Reporte):
Activa tu radar para detectar estos errores lógicos comunes:
Alerta Tipo 1 (Falso Positivo Físico): Avance físico de Actividad o MP > 0, pero Total Obligaciones = $0 (Sin justificación de gestión). Diagnóstico: Posible error de digitación o reporte sin soporte.
Alerta Tipo 2 (Omisión de Reporte Físico): Total Obligaciones altas (> 30%), pero Avance Físico = 0, y en las 'Dificultades', al reportar MP, NO explican que el proyecto está en etapa meramente contractual o precontractual. Esta situación es crítica para las ACTIVIDADES, pues la ejecución financiera debe ir acompañada de avance físico; en cambio, para los productos sí es posible que haya avance financiero y se reporte el avance físico en 0 pues no se ha consolidado la entrega del bien o servicio. Diagnóstico: Ejecutaron recursos pero olvidaron reportar el avance físico asociado.
Alerta Tipo 3 (Desconexión Jerárquica): Meta de Producto reporta avance muy alto (ej. 100%), pero el 'Promedio Avance Actividades' es críticamente bajo (ej. < 30%). Diagnóstico: Inconsistencia entre el proyecto (PA) y la meta (PI). Para el promedio de actividades es importante distinguir y promediar solo las actividades de 1 proyecto de inversión. Esta situación es un reporte inconsistente que debería tener alguna explicación al cotejar el reporte cualitativo de las actividades y el reporte cualitativo de la MP; por ejemplo, podría ser que una MP tiene 2 proyectos que contribuyen al cumplimiento, 1 de los cuales presenta avance coherente en sus actividades y productos, mientras que el otro no.

4. Calidad Narrativa y Veracidad:
Compara el número del "Resultado" contra el detalle del "Principal Logro". ¿La narrativa describe CÓMO se lograron las unidades reportadas? (Ej: Si el resultado dice 50, pero la narrativa solo describe 11, califica como "Narrativa Insuficiente").
Logro: Debe describir QUÉ se hizo para justificar el número de avance en el periodo evaluado. El logro principal se refiere a un texto cualitativo conciso que dé cuenta del principal logro alcanzado por la dependencia en el cumplimiento de la MP. No debe ser excesivamente detallado, pero SI debe comunicar lo hecho. Textos genéricos ("se avanzó según lo planeado") son causal de devolución.. Este logro resumido es el que se suele usar en los informes consolidados o estrategias de comunicaciones para informar a la ciudadanía sobre lo que se hace en la entidad.
Análisis: Debe detallar CÓMO y DÓNDE (municipios, grupos poblacionales). Textos genéricos ("se avanzó según lo planeado") son causal de devolución. Este apartado requiere mayor rigor técnico, pues la idea es que el enlace de SODR pueda leer y comprender con mayor grado de detalle en qué consiste en valor de avance reportado para la MP y cómo se interpreta ese valor. Por ejemplo, si la MP es de asistencias técnicas y se reportan 6 de 12 realizadas, se esperaría que haya una breve contextualización: dónde se realizaron las asistencias, con qué tipo de público o a qué entidades se enfocó, de qué temas se trataba, etc.
Dificultades: Especialmente en cortes trimestrales (donde el avance físico puede ser bajo), es OBLIGATORIO usar el campo de dificultades para explicar si los retrasos son normativos, contractuales o de planeación. Este campo es de apoyo para que la dependencia explique situaciones que afectan el cumplimiento.
Nota: Si una MP no tiene avance físico (reporte in 0), no debería tener reporte de principal logro o análisis de logro, sino de dificultades.
"""

    estructura_salida = """
ESTRUCTURA DE ANÁLISIS POR META (TU FLUJO DE PENSAMIENTO)
Para cada meta que analices, ejecuta mentalmente estas fases antes de emitir tu respuesta:
FASE 1: Mapeo Temporal y Semáforo de Desviación
FASE 2: Evaluación de Coherencia Integral (El Juicio)
FASE 3: Retroalimentación Dirigida (Feedback Técnico)

INSTRUCCIONES DE SALIDA (FORMATO ESTRICTO DE RESPUESTA)
Espera mi instrucción para procesar cada bloque de metas. Tu respuesta por cada meta debe seguir estrictamente este formato Markdown:

🔎 Revisión Técnica: [CÓDIGO DE LA META]
Descripción de Meta
Comportamiento del Indicador
1. Semáforo de Consistencia (Corte: [Periodo]):
Meta Vigencia: [Valor] | Avance Reportado: [Valor] | % Avance: [Cálculo%]
Avance Promedio Actividades (PA): [X%]
Ejecución Financiera (Obligaciones): $[Valor]
Estado: [🟢 CONSISTENTE / 🟡 ALERTA DE REVISIÓN / 🔴 INCONSISTENCIA CRÍTICA]

2. Análisis y Sistema de Alertas:
[Lista aquí los hallazgos técnicos derivados de la Fase 2. Utiliza las tipologías de alertas definidas. Sé duro y directo.]

3. Veredicto y Retroalimentación:
Dictamen Sugerido: [APROBADA] o [DEVUELTA]
Feedback Técnico para la Entidad (Leer críticamente para enviar como observación):
Redacta aquí un párrafo formal, institucional y respetuoso dirigido al responsable. Debe contener:
1. Identificación clara del error o vacío técnico.
2. Requerimiento específico para subsanar el reporte in EVAPLAN.

Ten en cuenta: Cuida la precisión de la terminología usada, por ejemplo: Las obligaciones financieras son equivalentes a ejecución financiera, sin embargo, no son lo mismo que "Presupuesto Comprometido".

INSTRUCCIÓN DE INICIO:
Si has asimilado todas estas reglas, comprendes la importancia de la temporalidad, y estás listo para aplicar el Sistema Integrado de Alertas y la evaluación narrativa cruzada, responde ÚNICAMENTE con el siguiente texto:
"Entendido. Soy BOT_SODR_EVAPLAN, tu auditor técnico experto. He configurado la temporalidad y el sistema de alertas. Por favor, indícame el Periodo de Corte y carga los datos de las metas o el archivo integrado para iniciar la auditoría rigurosa."
"""

    if periodo == "Revisión acumulada de primer trimestre":
        bloque_config = """
[BLOQUE DE CONFIGURACIÓN DE LA REVISIÓN]
Nota para el GEM: El usuario te indica que el periodo de revisión corresponde al Primer Trimestre de la vigencia. Adapta tu juicio a esta temporalidad.
Periodo de Corte Actual: Primer Trimestre de 2026 (Q1 2026).

Lógica de Temporalidad: Al ser un reporte trimestral parcial, no se exige el 100% del cumplimiento final de la meta anual. Se evalúa que el avance reportado (físico y financiero) sea coherente con los primeros meses del año. Un reporte de 100% en Q1 debe ser revisado con extrema lupa, y un reporte de 0% con alta ejecución financiera requiere justificación de etapa precontractual.
Aclaración: Si en reportes parciales se registra 100% de avance, la descripción cualitativa debe ayudar a entender cómo se consiguió ese nivel de avance. También se debería aclarar que la meta busca un sostenimiento o continuidad en el producto (bien o servicio) que está entregando, de modo que, si ya alcanzó el 100% en Q1 (por ejemplo), pues se va a mantener ese nivel de entrega; o si por el contrario, ya no se va a entregar nada más. 
"""
    elif periodo == "Revisión acumulada del Primer semestre":
        bloque_config = """
[BLOQUE DE CONFIGURACIÓN DE LA REVISIÓN]
Nota para el GEM: El usuario te indica que el periodo de revisión corresponde al Primer Semestre acumulado. Adapta tu juicio a esta temporalidad de mitad de año.
Periodo de Corte Actual: Primer Semestre de 2026 (Q2 2026).

Lógica de Temporalidad: Al ser un reporte acumulado a mitad de año (Corte a Junio), se espera una ejecución física cercana al 40%-50% o una justificación contractual clara si es menor. Para los casos donde en Q2 ya se alcanzó el valor esperado del 100% anual y ya no se va a entregar más producto, es un dato crítico porque supone que ya no se debería ejecutar más recurso (contratar para la ejecución de actividades ligadas a esa meta) a través de esa MP. Este es un dato relevante que debe hacerse notar para que lo tengan en cuenta los enlaces SODR.
"""
    elif periodo == "Revisión Acumulada de Tercer Semestre":
        bloque_config = """
[BLOQUE DE CONFIGURACIÓN DE LA REVISIÓN]
Nota para el GEM: El periodo de revisión corresponde a la revisión acumulada de Tercer Semestre (Periodo extendido multianual o ajuste de ciclo). 
"""
    elif periodo == "Revisión Acumulada y Proyectada a Cierre de Vigencia":
        bloque_config = """
[BLOQUE DE CONFIGURACIÓN DE LA REVISIÓN]
Nota para el GEM: El periodo de revisión pferece al precierre de la vigencia, analizando la ejecución real frente a proyecciones de cierre.
"""
    else:
        bloque_config = """
[BLOQUE DE CONFIGURACIÓN DE LA REVISIÓN]
Nota para el GEM: El periodo de revisión corresponde al Cierre Final de la Vigencia. El juicio aquí es definitivo y estricto frente a metas anuales al 100%.
"""

    return perfil_mision + bloque_config + contexto_usuario + reglas_oro + estructura_salida


# ============================================================
# FUNCIÓN LIMPIEZA PRESUPUESTAL
# ============================================================
def limpiar_moneda(serie):
    return pd.to_numeric(
        (
            serie.astype(str)
            .str.replace(r"\$", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        ),
        errors="coerce"
    )

# ============================================================
# FUNCIÓN PDF CON NEGRILLA
# ============================================================
def escribir_bloque(c, texto, y, width, height, margen_x, margen_y, tamaño=10, negrilla=False):
    textobject = c.beginText(margen_x, y)
    fuente = "Helvetica-Bold" if negrilla else "Helvetica"
    textobject.setFont(fuente, tamaño)

    for linea in wrap(str(texto), 95):
        if textobject.getY() <= margen_y:
            c.drawText(textobject)
            c.showPage()
            textobject = c.beginText(margen_x, height - margen_y)
            textobject.setFont(fuente, tamaño)
        textobject.textLine(linea)

    c.drawText(textobject)
    return textobject.getY() - 12

columnas_base_pi = [
    'Código de Meta', 'Descripción de Meta', 'Comportamiento del Indicador',
    'Valor Proyectado', 'Resultado', '2026',
    'Principal Logro en Función del Cumplimiento', 'Análisis del Logro',
    'Dificultades o Gestiones'
]

columnas_focalizacion = [
    'Negro, Mulato, Afrodescendiente, Raizal y Palenquero', 'Indígena', 'Room',
    'Campesinos', 'Niños Niñas y Adolescentes', 'Primera Infancia', 'Juventud',
    'Personas Mayores', 'Mujer', 'LGTBIQ+', 'Personas con Discapacidad y sus Curadores',
    'Personas Vulnerables', 'Habitantes de o en Calle', 'Víctimas de Violencia de Género',
    'Víctimas del Conflicto', 'Reincorporados', 'Comunales', 'Interreligioso',
    'Rescatistas de Animales', 'Migrantes', 'Retornados', 'Otros', '¿Cuál Otro?'
]

columnas_mp = [
    "Cód. MP", "Descripción MP", "Cód. Proyecto", "Nombre Proyecto",
    "Ppto. Total Obligaciones", "Ppto. Definitivo", "% Avance x Actividad"
]

def consolidar_proyectos(grupo):
    proyectos = []
    pares_unicos = set()
    for _, fila in grupo.iterrows():
        codigo = str(fila["Cód. Proyecto"]).strip()
        nombre = str(fila["Nombre Proyecto"]).strip()
        if codigo == "" or codigo.lower() == "nan" or nombre == "" or nombre.lower() == "nan":
            continue
        par = f"{codigo} - {nombre}"
        if par not in pares_unicos:
            pares_unicos.add(par)
            proyectos.append(par)
    return " | ".join(proyectos)

# ============================================================
# INTERFAZ DE USUARIO - CARGA DE ARCHIVOS CON PERSISTENCIA
# ============================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plan Indicativo")
    # Al asignarle una key única, el componente recuerda el archivo de forma nativa en st.session_state
    file_pi = st.file_uploader("Subir 'Informe de Plan Indicativo MP.xlsx'", type=["xlsx"], key="file_pi_uploader")

with col2:
    st.subheader("2. Plan de Acción / Centralizadas")
    file_pa = st.file_uploader("Subir 'Centralizadas.xlsx'", type=["xlsx"], key="file_pa_uploader")

if file_pi and file_pa:
    st.success("¡Ambos archivos cargados con éxito! Presiona el botón para procesar.")
    
    if st.button("🚀 Procesar Datos e Integrar", type="primary"):
        try:
            with st.spinner("Procesando información presupuestal y cualitativa..."):
                df_pi = pd.read_excel(file_pi, sheet_name="Sheet1", header=1)
                df_pi.columns = df_pi.columns.astype(str).str.strip()

                faltantes_pi = [col for col in columnas_base_pi if col not in df_pi.columns]
                if faltantes_pi:
                    st.error(f"Faltan columnas obligatorias en el Plan Indicativo: {faltantes_pi}")
                    st.stop()

                columnas_focalizacion_activas = []
                for col in columnas_focalizacion:
                    if col in df_pi.columns:
                        tiene_datos = False
                        for valor in df_pi[col]:
                            if pd.isna(valor):
                                continue
                            try:
                                if float(valor) != 0:
                                    tiene_datos = True
                                    break
                            except:
                                valor_texto = str(valor).strip()
                                if valor_texto != "" and valor_texto.lower() != "nan" and valor_texto != "0":
                                    tiene_datos = True
                                    break
                        if tiene_datos:
                            columnas_focalizacion_activas.append(col)

                columnas_pi_finales = columnas_base_pi + columnas_focalizacion_activas
                df_pi = df_pi[columnas_pi_finales].copy()

                for col in ['Resultado', '2026']:
                    df_pi[col] = pd.to_numeric(df_pi[col], errors='coerce')
                df_pi['Relacion_Resultado_MP_vs_2026'] = df_pi['Resultado'] / df_pi['2026']

                df_pa = pd.read_excel(file_pa, sheet_name="Sheet1", header=1)
                df_pa.columns = df_pa.columns.astype(str).str.strip()

                faltantes_mp = [col for col in columnas_mp if col not in df_pa.columns]
                if faltantes_mp:
                    st.error(f"Faltan columnas obligatorias en Centralizadas: {faltantes_mp}")
                    st.stop()

                df_mp = df_pa[columnas_mp].copy()

                for col in ["Ppto. Total Obligaciones", "Ppto. Definitivo"]:
                    df_mp[col] = limpiar_moneda(df_mp[col])

                df_mp["% Avance x Actividad"] = pd.to_numeric(df_mp["% Avance x Actividad"], errors="coerce")
                df_mp["Avance_Actividad_01"] = df_mp["% Avance x Actividad"] / 100

                df_agrupado = (
                    df_mp.groupby(["Cód. MP", "Descripción MP"], as_index=False)
                    .apply(lambda grupo: pd.Series({
                        "Proyectos_Asociados": consolidar_proyectos(grupo),
                        "Suma_Ppto_Total_Obligaciones": grupo["Ppto. Total Obligaciones"].sum(),
                        "Suma_Ppto_Definitivo": grupo["Ppto. Definitivo"].sum(),
                        "Promedio_Avance_Actividades_01": grupo["Avance_Actividad_01"].mean()
                    }))
                    .reset_index(drop=True)
                )

                df_agrupado["Relacion_Obligaciones_vs_Definitivo"] = (
                    df_agrupado["Suma_Ppto_Total_Obligaciones"] / df_agrupado["Suma_Ppto_Definitivo"]
                )

                df_integrado = df_pi.merge(df_agrupado, left_on="Código de Meta", right_on="Cód. MP", how="left")
                if "Cód. MP" in df_integrado.columns:
                    df_integrado = df_integrado.drop(columns=["Cód. MP"])

                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                    df_integrado.to_excel(writer, index=False, sheet_name="MP_PI_PA")
                
                # Almacenamos el resultado del procesamiento en el session_state
                st.session_state["excel_data"] = output_excel.getvalue()

                output_pdf = io.BytesIO()
                c = canvas.Canvas(output_pdf, pagesize=LETTER)
                width, height = LETTER
                margen_x, margen_y = 50, 50

                def obtener_focalizacion(row):
                    focalizaciones = []
                    for col in columnas_focalizacion_activas:
                        valor = row.get(col, None)
                        if pd.isna(valor):
                            continue
                        try:
                            if float(valor) != 0:
                                focalizaciones.append(f"{col}: {valor}")
                        except:
                            valor_texto = str(valor).strip()
                            if valor_texto != "" and valor_texto.lower() != "nan" and valor_texto != "0":
                                focalizaciones.append(f"{col}: {valor_texto}")
                    return " | ".join(focalizaciones) if focalizaciones else "No reporta focalización."

                for _, row in df_integrado.iterrows():
                    y = height - margen_y
                    y = escribir_bloque(c, "META PRODUCTO", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, f"Código de Meta: {row.get('Código de Meta','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Descripción de Meta: {row.get('Descripción de Meta','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Descripción MP: {row.get('Descripción MP','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Comportamiento del Indicador: {row.get('Comportamiento del Indicador','')}", y, width, height, margen_x, margen_y)
                    y -= 8

                    y = escribir_bloque(c, "PLAN INDICATIVO (PI)", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, f"Valor Proyectado: {row.get('Valor Proyectado','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Resultado: {row.get('Resultado','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Programación 2026: {row.get('2026','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Relación Resultado / 2026: {row.get('Relacion_Resultado_MP_vs_2026','')}", y, width, height, margen_x, margen_y)
                    y -= 8

                    y = escribir_bloque(c, "PROYECTOS ASOCIADOS", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, f"Proyectos Asociados: {row.get('Proyectos_Asociados','')}", y, width, height, margen_x, margen_y)
                    y -= 8

                    y = escribir_bloque(c, "FOCALIZACIÓN", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, obtener_focalizacion(row), y, width, height, margen_x, margen_y)
                    y -= 8

                    y = escribir_bloque(c, "PLAN DE ACCIÓN (PA)", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, f"Suma Presupuesto Definitivo: {row.get('Suma_Ppto_Definitivo','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Suma Total Obligaciones: {row.get('Suma_Ppto_Total_Obligaciones','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Promedio Avance Actividades: {row.get('Promedio_Avance_Actividades_01','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Relación Obligaciones / Definitivo: {row.get('Relacion_Obligaciones_vs_Definitivo','')}", y, width, height, margen_x, margen_y)
                    y -= 8

                    y = escribir_bloque(c, "ANÁLISIS CUALITATIVO", y, width, height, margen_x, margen_y, 11, True)
                    y = escribir_bloque(c, f"Principal Logro en Función del Cumplimiento: {row.get('Principal Logro en Función del Cumplimiento','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Análisis del Logro: {row.get('Análisis del Logro','')}", y, width, height, margen_x, margen_y)
                    y = escribir_bloque(c, f"Dificultades o Gestiones: {row.get('Dificultades o Gestiones','')}", y, width, height, margen_x, margen_y)
                    c.showPage()

                c.save()
                st.session_state["pdf_data"] = output_pdf.getvalue()
                st.session_state["prompt_final"] = generar_prompt_sistema(periodo_seleccionado)
                st.session_state["procesado_exitoso"] = True
                
                st.balloons()

        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {e}")
            st.session_state["procesado_exitoso"] = False

# ============================================================
# RENDERIZADO PERSISTENTE DE RESULTADOS
# ============================================================
# Esto garantiza que los botones y prompts sigan visibles aunque el usuario cambie de página o interactúe
if st.session_state["procesado_exitoso"]:
    st.success("🎉 ¡Proceso finalizado con éxito!")

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.download_button(
            label="📥 Descargar Matriz Integrada (Excel)",
            data=st.session_state["excel_data"],
            file_name="MP_PI_PA_Integrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with d_col2:
        st.download_button(
            label="📥 Descargar Reporte Completo (PDF)",
            data=st.session_state["pdf_data"],
            file_name="MP_PI_PA_Gemini_Completo.pdf",
            mime="application/pdf"
        )

    st.markdown("---")
    st.subheader("🤖 Asistente de Auditoría EVAPLAN (Prompt Listo)")
    
    # Recalcula el prompt dinámicamente si el usuario cambia el periodo en la barra lateral sin volver a procesar archivos
    prompt_dinamico = generar_prompt_sistema(periodo_seleccionado)
    
    with st.expander("📋 Ver y Copiar Propuesta de Prompt para Gemini/ChatGPT", expanded=True):
        st.text_area(
            label="Puedes copiar el texto completo usando el botón superior derecho:",
            value=prompt_dinamico,
            height=350
        )
else:
    if not (file_pi and file_pa):
        st.info("💡 Por favor, sube ambos archivos de Excel para habilitar la unificación de los planes.")
