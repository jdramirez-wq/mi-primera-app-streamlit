import streamlit as st
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from textwrap import wrap
import io

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Consolidador PI + PA + PDF",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Consolidación Plan Indicativo (PI) + Plan de Acción (PA)")
st.write("Sube los archivos de Excel correspondientes para procesar, consolidar y descargar los resultados en formatos Excel y PDF.")

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

# Columnas predefinidas
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
# INTERFAZ DE USUARIO - CARGA DE ARCHIVOS
# ============================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plan Indicativo")
    file_pi = st.file_uploader("Subir 'Informe de Plan Indicativo MP.xlsx'", type=["xlsx"])

with col2:
    st.subheader("2. Plan de Acción / Centralizadas")
    file_pa = st.file_uploader("Subir 'Centralizadas.xlsx'", type=["xlsx"])

if file_pi and file_pa:
    st.success("¡Ambos archivos cargados con éxito! Presiona el botón para procesar.")
    
    if st.button("🚀 Procesar Datos e Integrar", type="primary"):
        try:
            with st.spinner("Procesando información presupuestal y cualitativa..."):
                # Lectura de datos
                df_pi = pd.read_excel(file_pi, sheet_name="Sheet1", header=1)
                df_pi.columns = df_pi.columns.astype(str).str.strip()

                # Validación PI
                faltantes_pi = [col for col in columnas_base_pi if col not in df_pi.columns]
                if faltantes_pi:
                    st.error(f"Faltan columnas obligatorias en el Plan Indicativo: {faltantes_pi}")
                    st.stop()

                # Detección dinámica de focalización
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

                # Filtrado PI
                columnas_pi_finales = columnas_base_pi + columnas_focalizacion_activas
                df_pi = df_pi[columnas_pi_finales].copy()

                for col in ['Resultado', '2026']:
                    df_pi[col] = pd.to_numeric(df_pi[col], errors='coerce')
                df_pi['Relacion_Resultado_MP_vs_2026'] = df_pi['Resultado'] / df_pi['2026']

                # Lectura Plan de Acción
                df_pa = pd.read_excel(file_pa, sheet_name="Sheet1", header=1)
                df_pa.columns = df_pa.columns.astype(str).str.strip()

                # Validación PA
                faltantes_mp = [col for col in columnas_mp if col not in df_pa.columns]
                if faltantes_mp:
                    st.error(f"Faltan columnas obligatorias en Centralizadas: {faltantes_mp}")
                    st.stop()

                df_mp = df_pa[columnas_mp].copy()

                for col in ["Ppto. Total Obligaciones", "Ppto. Definitivo"]:
                    df_mp[col] = limpiar_moneda(df_mp[col])

                df_mp["% Avance x Actividad"] = pd.to_numeric(df_mp["% Avance x Actividad"], errors="coerce")
                df_mp["Avance_Actividad_01"] = df_mp["% Avance x Actividad"] / 100

                # Agrupación por MP
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

                # Integración final
                df_integrado = df_pi.merge(df_agrupado, left_on="Código de Meta", right_on="Cód. MP", how="left")
                if "Cód. MP" in df_integrado.columns:
                    df_integrado = df_integrado.drop(columns=["Cód. MP"])

                # ============================================================
                # CREACIÓN DE DESCARGAS EN MEMORIA (BytesIO)
                # ============================================================
                
                # 1. Archivo Excel
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                    df_integrado.to_excel(writer, index=False, sheet_name="MP_PI_PA")
                excel_data = output_excel.getvalue()

                # 2. Archivo PDF
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
                pdf_data = output_pdf.getvalue()

            st.balloons()
            st.success("🎉 ¡Proceso finalizado con éxito! Descarga tus reportes aquí abajo:")

            # Botones de descarga organizados en columnas
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.download_button(
                    label="📥 Descargar Matriz Integrada (Excel)",
                    data=excel_data,
                    file_name="MP_PI_PA_Integrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with d_col2:
                st.download_button(
                    label="📥 Descargar Reporte Completo (PDF)",
                    data=pdf_data,
                    file_name="MP_PI_PA_Gemini_Completo.pdf",
                    mime="application/pdf"
                )

            # Vista previa opcional en pantalla
            st.subheader("👀 Vista previa de los datos unificados (Primeras 5 filas)")
            st.dataframe(df_integrado.head())

        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("💡 Por favor, sube ambos archivos de Excel para habilitar la unificación de los planes.")
