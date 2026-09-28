import streamlit as st
import pandas as pd
from docx import Document
import io
import os
import zipfile

st.set_page_config(page_title="Generador Plan Concertado - SENA", layout="wide")

st.title("📋 Generador de Plan Concertado de Trabajo (SENA)")
st.write("Generador simplificado de Plan Concertado de Trabajo.")

# Ruta de la plantilla en el repositorio
PLANTILLA_PATH = "Plan de trabajo .docx"

if not os.path.exists(PLANTILLA_PATH):
    st.error(f"❌ No se encontró el archivo '{PLANTILLA_PATH}' en el repositorio.")
    st.stop()

# 1. Cargar archivo Excel (Juicios Evaluativos)
excel_file = st.file_uploader("1. Cargar archivo Excel de Juicios Evaluativos", type=["xlsx", "xls"])

if excel_file:
    try:
        # Fila 13 contiene los encabezados reales de SofiaPlus
        df = pd.read_excel(excel_file, header=12)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

    df.columns = [str(col).strip() for col in df.columns]
    cols_lista = list(df.columns)

    st.subheader("📌 1. Selección de Competencia y Resultado de Aprendizaje")
    col_comp, col_rap = st.columns(2)

    with col_comp:
        col_competencia = st.selectbox("Seleccione la columna de Competencia:", cols_lista)
        opciones_competencia = df[col_competencia].dropna().unique().tolist()
        competencia_seleccionada = st.selectbox("Competencia escogida:", opciones_competencia) if opciones_competencia else ""

    with col_rap:
        col_resultado = st.selectbox("Seleccione la columna de Resultado de Aprendizaje:", cols_lista)
        opciones_rap = df[col_resultado].dropna().unique().tolist()
        resultado_seleccionado = st.selectbox("Resultado de Aprendizaje escogido:", opciones_rap) if opciones_rap else ""

    st.subheader("⚙️ 2. Datos Generales del Instructor y Entrega")
    col_in1, col_in2, col_in3 = st.columns(3)

    with col_in1:
        instructor = st.text_input("Instructor", value="WILLIAM SANTIAGO HURTADO CARMONA")
        proyecto_formativo = st.text_input("Proyecto Formativo")
        fase_proyecto = st.text_input("Fase del Proyecto")

    with col_in2:
        forma_entrega = st.selectbox("Forma de Entrega", ["Digital", "Física"])
        tipo_reporte = st.radio("Tipo de Reporte", ["Inicial (En Blanco)", "Final (Con Marcación SI)"])

    with col_in3:
        fecha_inicial = st.date_input("Fecha Inicial")
        fecha_final = st.date_input("Fecha Final")

    st.subheader("👤 3. Mapeo de Aprendices")
    col_a1, col_a2, col_a3, col_a4 = st.columns(4)

    def buscar_index(patrones):
        for pat in patrones:
            for idx, col in enumerate(cols_lista):
                if pat.lower() in str(col).lower():
                    return idx
        return 0

    with col_a1:
        col_nom = st.selectbox("Columna Nombres", cols_lista, index=buscar_index(["nombre"]))
    with col_a2:
        col_ape = st.selectbox("Columna Apellidos", cols_lista, index=buscar_index(["apellido"]))
    with col_a3:
        col_tipo = st.selectbox("Columna Tipo Documento", cols_lista, index=buscar_index(["tipo"]))
    with col_a4:
        col_doc = st.selectbox("Columna N° Documento", cols_lista, index=buscar_index(["documento", "numero", "identificac"]))

    # Crear columna concatenada de Aprendices (Nombre + Apellido + Documento)
    df["Aprendiz_Concatenado"] = (
        df[col_nom].astype(str).str.strip() + " " + 
        df[col_ape].astype(str).str.strip() + " (" + 
        df[col_tipo].astype(str).str.strip() + " " + 
        df[col_doc].astype(str).str.strip() + ")"
    )

    # Filtrar registros válidos sin duplicados
    df_aprendices = df.drop_duplicates(subset=[col_doc]).dropna(subset=[col_nom]).reset_index(drop=True)

    st.markdown("---")
    st.subheader("🎯 Selección de Aprendices a Generar")

    generar_todos = st.checkbox("✅ Generar reporte para TODOS los aprendices", value=True)

    if not generar_todos:
        lista_opciones = df_aprendices["Aprendiz_Concatenado"].tolist()
        seleccionados = st.multiselect("Selecciona los aprendices:", options=lista_opciones)
        df_final = df_aprendices[df_aprendices["Aprendiz_Concatenado"].isin(seleccionados)]
    else:
        df_final = df_aprendices

    st.info(f"Se generarán **{len(df_final)}** documento(s).")

    def reemplazar_en_tabla(doc, reemplazos):
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for k, v in reemplazos.items():
                            if k in paragraph.text:
                                paragraph.text = paragraph.text.replace(k, str(v))

    if st.button("🚀 Generar Todos los Planes Concertados"):
        if df_final.empty:
            st.warning("Debe haber al menos un aprendiz para procesar.")
            st.stop()

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for _, row in df_final.iterrows():
                doc = Document(PLANTILLA_PATH)

                # Reemplazos directos
                reemplazos = {
                    "«Nombre»": str(row[col_nom]),
                    "«Apellidos»": str(row[col_ape]),
                    "«Tipo_de_Doc»": str(row[col_tipo]),
                    "«N_Documento»": str(row[col_doc]),
                    "JHON CUENTAS DE CARO": instructor,
                    "WILLIAM SANTIAGO HURTADO CARMONA": instructor
                }

                # Configuración Actividades 1 a 10 (Inicial vs Final)
                for i in range(1, 11):
                    tag_si = f"«Act_{i}_Si»"
                    tag_no = f"«Act_{i}_No»"
                    if tipo_reporte.startswith("Inicial"):
                        reemplazos[tag_si] = ""
                        reemplazos[tag_no] = ""
                    else:
                        reemplazos[tag_si] = "SI"
                        reemplazos[tag_no] = ""

                reemplazar_en_tabla(doc, reemplazos)

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                nombre_doc = f"Plan_Concertado_{row[col_nom]}_{row[col_ape]}.docx".replace(" ", "_")
                zip_file.writestr(nombre_doc, doc_buffer.getvalue())

        zip_buffer.seek(0)

        st.success(f"¡Se generaron con éxito **{len(df_final)}** planes concertados!")
        st.download_button(
            label="📥 Descargar Paquete Completo (.ZIP)",
            data=zip_buffer,
            file_name="Planes_Concertados_Todos.zip",
            mime="application/zip"
        )
