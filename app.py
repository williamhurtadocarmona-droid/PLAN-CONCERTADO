import streamlit as st
import pandas as pd
from docx import Document
import io
import os

st.set_page_config(page_title="Generador Plan Concertado - SENA", layout="wide")

st.title("📋 Generador de Plan Concertado de Trabajo (SENA)")
st.write("Carga el archivo Excel de **Juicios Evaluativos** para generar el documento Word.")

# Ruta de la plantilla incluida en el repositorio de GitHub
PLANTILLA_PATH = "Plan de trabajo .docx"

# Verificar si la plantilla existe en el repositorio
if not os.path.exists(PLANTILLA_PATH):
    st.error(f"❌ No se encontró el archivo '{PLANTILLA_PATH}' en la raíz del repositorio.")
    st.stop()

# 1. Carga únicamente del archivo Excel
excel_file = st.file_uploader("1. Cargar archivo Excel (Juicios Evaluativos)", type=["xlsx", "xls"])

if excel_file:
    # Leer el archivo de Excel
    df = pd.read_excel(excel_file)
    
    st.subheader("⚙️ Configuración del Documento")
    
    col_input1, col_input2, col_input3 = st.columns(3)
    
    with col_input1:
        instructor = st.text_input("Nombre del Instructor", value="JHON CUENTAS DE CARO")
        proyecto_formativo = st.text_input("Proyecto Formativo")
        fase_proyecto = st.text_input("Fase del Proyecto")
        
    with col_input2:
        forma_entrega = st.selectbox("Forma de Entrega de Actividad", ["Digital", "Física"])
        tipo_documento = st.radio("Tipo de Reporte a Generar", ["Inicial (Sin Marcar Entrega)", "Final (Con SI/NO)"])
        
    with col_input3:
        fecha_inicial = st.date_input("Fecha Inicial (Concertada)")
        fecha_final = st.date_input("Fecha Final de Entrega")
        observaciones = st.text_area("Observaciones", value="")

    st.subheader("👤 Selección de Aprendiz")
    
    col_sel1, col_sel2, col_sel3, col_sel4 = st.columns(4)
    
    with col_sel1:
        col_nombre = st.selectbox("Columna 'Nombre'", df.columns, index=0 if "Nombre" in df.columns else 0)
    with col_sel2:
        col_apellido = st.selectbox("Columna 'Apellidos'", df.columns, index=1 if "Apellidos" in df.columns else 0)
    with col_sel3:
        col_tipo_doc = st.selectbox("Columna 'Tipo de Documento'", df.columns, index=2 if "Tipo_de_Doc" in df.columns else 0)
    with col_sel4:
        col_doc = st.selectbox("Columna 'N° Documento'", df.columns, index=3 if "N_Documento" in df.columns else 0)

    # Menú desplegable para elegir el aprendiz
    aprendices_lista = df.apply(lambda row: f"{row[col_nombre]} {row[col_apellido]} ({row[col_doc]})", axis=1).tolist()
    aprendiz_seleccionado_idx = st.selectbox("Seleccione el Aprendiz:", range(len(aprendices_lista)), format_func=lambda x: aprendices_lista[x])

    def reemplazar_en_tabla(doc, reemplazos):
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for k, v in reemplazos.items():
                            if k in paragraph.text:
                                paragraph.text = paragraph.text.replace(k, str(v))

    if st.button("🚀 Generar Plan Concertado"):
        row = df.iloc[aprendiz_seleccionado_idx]
        
        # Cargar automáticamente la plantilla interna
        doc = Document(PLANTILLA_PATH)
        
        # Mapa de reemplazo de campos
        reemplazos = {
            "«Nombre»": str(row[col_nombre]),
            "«Apellidos»": str(row[col_apellido]),
            "«Tipo_de_Doc»": str(row[col_tipo_doc]),
            "«N_Documento»": str(row[col_doc]),
            "JHON CUENTAS DE CARO": instructor,
        }
        
        # Lógica para la sección de marcación ¿Entrego la Actividad?
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
