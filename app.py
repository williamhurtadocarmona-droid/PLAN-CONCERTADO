import streamlit as st
import pandas as pd
from docx import Document
import io
import os
import zipfile

st.set_page_config(page_title="Generador Plan Concertado - SENA", layout="wide")

st.title("📋 Generador de Plan Concertado de Trabajo (SENA)")
st.write("Carga el archivo Excel de **Juicios Evaluativos** para generar los documentos Word.")

# Ruta de la plantilla en el repositorio
PLANTILLA_PATH = "Plan de trabajo .docx"

if not os.path.exists(PLANTILLA_PATH):
    st.error(f"❌ No se encontró el archivo '{PLANTILLA_PATH}' en la raíz del repositorio.")
    st.stop()

# 1. Carga del archivo Excel
excel_file = st.file_uploader("1. Cargar archivo Excel (Juicios Evaluativos)", type=["xlsx", "xls"])

if excel_file:
    try:
        # Se indica header=12 porque la fila 13 contiene los encabezados reales de SofiaPlus/SENA
        df = pd.read_excel(excel_file, header=12)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

    # Limpiar nombres de columnas y convertir a string
    df.columns = [str(col).strip() for col in df.columns]

    st.subheader("⚙️ Configuración del Documento")
    
    col_input1, col_input2, col_input3 = st.columns(3)
    
    with col_input1:
        instructor = st.text_input("Nombre del Instructor", value="WILLIAM SANTIAGO HURTADO CARMONA")
        proyecto_formativo = st.text_input("Proyecto Formativo")
        fase_proyecto = st.text_input("Fase del Proyecto")
        
    with col_input2:
        forma_entrega = st.selectbox("Forma de Entrega de Actividad", ["Digital", "Física"])
        tipo_documento = st.radio("Tipo de Reporte a Generar", ["Inicial (Sin Marcar Entrega)", "Final (Con SI/NO)"])
        
    with col_input3:
        fecha_inicial = st.date_input("Fecha Inicial (Concertada)")
        fecha_final = st.date_input("Fecha Final de Entrega")
        observaciones = st.text_area("Observaciones", value="")

    st.subheader("👤 Selección y Mapeo de Columnas")
    
    cols_lista = list(df.columns)
    
    # Función corregida para buscar columnas sin error de AttributeError
    def buscar_columna(patrones, lista_cols):
        for pat in patrones:
            for idx, col in enumerate(lista_cols):
                if pat.lower() in str(col).lower():
                    return idx
        return 0

    idx_nom = buscar_columna(["nombre"], cols_lista)
    idx_ape = buscar_columna(["apellido"], cols_lista)
    idx_tipo = buscar_columna(["tipo document", "tipo_doc", "tipo de doc"], cols_lista)
    idx_doc = buscar_columna(["numero document", "n_documento", "documento", "identificac"], cols_lista)
    idx_estado = buscar_columna(["estado"], cols_lista)

    col_sel1, col_sel2, col_sel3, col_sel4 = st.columns(4)

    with col_sel1:
        col_nombre = st.selectbox("Columna 'Nombre'", cols_lista, index=idx_nom)
    with col_sel2:
        col_apellido = st.selectbox("Columna 'Apellidos'", cols_lista, index=idx_ape)
    with col_sel3:
        col_tipo_doc = st.selectbox("Columna 'Tipo de Documento'", cols_lista, index=idx_tipo)
    with col_sel4:
        col_doc = st.selectbox("Columna 'N° Documento'", cols_lista, index=idx_doc)

    # Filtrado por Estado "EN FORMACIÓN"
    st.markdown("---")
    st.subheader("🎯 Filtrado y Selección de Aprendices")
    
    col_filt1, col_filt2 = st.columns(2)
    with col_filt1:
        col_estado = st.selectbox("Columna de Estado del Aprendiz", cols_lista, index=idx_estado)
    
    with col_filt2:
        filtrar_en_formacion = st.checkbox("Filtrar solo aprendices 'EN FORMACION'", value=True)

    if filtrar_en_formacion and col_estado in df.columns:
        df_filtrado = df[df[col_estado].astype(str).str.upper().str.contains("FORMACI", na=False)].copy()
        st.info(f"Se filtraron **{len(df_filtrado)}** aprendices en estado 'EN FORMACIÓN' de un total de {len(df)} registro(s).")
    else:
        df_filtrado = df.copy()

    # Eliminar duplicados por número de documento si los hay
    df_filtrado = df_filtrado.drop_duplicates(subset=[col_doc]).reset_index(drop=True)

    # Opción para seleccionar TODOS
    seleccionar_todos = st.checkbox("✅ Seleccionar TODOS los aprendices filtrados", value=False)
    
    lista_opciones = [
        f"{row[col_nombre]} {row[col_apellido]} ({row[col_doc]})" 
        for _, row in df_filtrado.iterrows()
    ]
    
    if seleccionar_todos:
        aprendices_seleccionados_indices = list(range(len(df_filtrado)))
        st.success(f"Se procesarán los **{len(df_filtrado)}** aprendices seleccionados.")
    else:
        indices_elegidos = st.multiselect(
            "Seleccione uno o varios Aprendices:",
            options=list(range(len(lista_opciones))),
            format_func=lambda x: lista_opciones[x]
        )
        aprendices_seleccionados_indices = indices_elegidos

    def reemplazar_en_tabla(doc, reemplazos):
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for k, v in reemplazos.items():
                            if k in paragraph.text:
                                paragraph.text = paragraph.text.replace(k, str(v))

    if st.button("🚀 Generar Plan(es) Concertado(s)"):
        if not aprendices_seleccionados_indices:
            st.warning("⚠️ Debe seleccionar al menos un aprendiz.")
            st.stop()

        # Caso 1: Generar 1 solo aprendiz -> Descargar archivo Word
        if len(aprendices_seleccionados_indices) == 1:
            idx = aprendices_seleccionados_indices[0]
            row = df_filtrado.iloc[idx]
            
            doc = Document(PLANTILLA_PATH)
            
            reemplazos = {
                "«Nombre»": str(row[col_nombre]),
                "«Apellidos»": str(row[col_apellido]),
                "«Tipo_de_Doc»": str(row[col_tipo_doc]),
                "«N_Documento»": str(row[col_doc]),
                "JHON CUENTAS DE CARO": instructor,
            }
            
            for i in range(1, 11):
                tag_si = f"«Act_{i}_Si»"
                tag_no = f"«Act_{i}_No»"
                if tipo_documento.startswith("Inicial"):
                    reemplazos[tag_si] = ""
                    reemplazos[tag_no] = ""
                else:
                    reemplazos[tag_si] = "X"
                    reemplazos[tag_no] = ""

            reemplazar_en_tabla(doc, reemplazos)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            nombre_archivo = f"Plan_Concertado_{row[col_nombre]}_{row[col_apellido]}.docx"
            
            st.success(f"¡Documento generado exitosamente para **{row[col_nombre]} {row[col_apellido]}**!")
            st.download_button(
                label="📥 Descargar Documento Word",
                data=buffer,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # Caso 2: Generar varios aprendices -> Descargar archivo ZIP
        else:
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx in aprendices_seleccionados_indices:
                    row = df_filtrado.iloc[idx]
                    doc = Document(PLANTILLA_PATH)
                    
                    reemplazos = {
                        "«Nombre»": str(row[col_nombre]),
                        "«Apellidos»": str(row[col_apellido]),
                        "«Tipo_de_Doc»": str(row[col_tipo_doc]),
                        "«N_Documento»": str(row[col_doc]),
                        "JHON CUENTAS DE CARO": instructor,
                    }
                    
                    for i in range(1, 11):
                        tag_si = f"«Act_{i}_Si»"
                        tag_no = f"«Act_{i}_No»"
                        if tipo_documento.startswith("Inicial"):
                            reemplazos[tag_si] = ""
                            reemplazos[tag_no] = ""
                        else:
                            reemplazos[tag_si] = "X"
                            reemplazos[tag_no] = ""

                    reemplazar_en_tabla(doc, reemplazos)
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    nombre_doc = f"Plan_Concertado_{row[col_nombre]}_{row[col_apellido]}.docx"
                    zip_file.writestr(nombre_doc, doc_buffer.getvalue())

            zip_buffer.seek(0)
            
            st.success(f"¡Se generaron con éxito **{len(aprendices_seleccionados_indices)}** documentos!")
            st.download_button(
                label="📥 Descargar TODOS los Planes (.ZIP)",
                data=zip_buffer,
                file_name="Planes_Concertados_Grupo.zip",
                mime="application/zip"
            )
