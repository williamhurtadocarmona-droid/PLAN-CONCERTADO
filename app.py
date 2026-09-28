import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Generador Plan Concertado - SENA", layout="wide")

st.title("📋 Generador de Plan Concertado de Trabajo (SENA)")
st.write("Carga el archivo Excel de **Juicios Evaluativos** y genera el documento Word individual o por grupo.")

# 1. Carga de archivos
col_file1, col_file2 = st.columns(2)

with col_file1:
    excel_file = st.file_uploader("1. Cargar archivo Excel (Juicios Evaluativos)", type=["xlsx", "xls"])

with col_file2:
    word_template_file = st.file_uploader("2. Cargar Plantilla Word (Plan de Trabajo .docx)", type=["docx"])

if excel_file and word_template_file:
    # Leer el archivo de Excel
    df = pd.read_excel(excel_file)
    
    st.subheader("⚙️ Configuración del Documento")
    
    # 2. Formulario para campos generales / manuales
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

    # 3. Selección del aprendiz desde los encabezados/columnas del Excel
    st.subheader("👤 Selección de Aprendiz")
    
    col_sel1, col_sel2, col_sel3 = st.columns(3)
    
    with col_sel1:
        col_nombre = st.selectbox("Columna 'Nombre'", df.columns, index=0 if "Nombre" in df.columns else 0)
    with col_sel2:
        col_apellido = st.selectbox("Columna 'Apellidos'", df.columns, index=1 if "Apellidos" in df.columns else 0)
    with col_sel3:
        col_doc = st.selectbox("Columna 'N° Documento'", df.columns, index=2 if "N_Documento" in df.columns else 0)

    col_tipo_doc = st.selectbox("Columna 'Tipo de Documento'", df.columns, index=3 if "Tipo_de_Doc" in df.columns else 0)

    # Menú desplegable para elegir un aprendiz en particular
    aprendices_lista = df.apply(lambda row: f"{row[col_nombre]} {row[col_apellido]} ({row[col_doc]})", axis=1).tolist()
    aprendiz_seleccionado_idx = st.selectbox("Seleccione el Aprendiz:", range(len(aprendices_lista)), format_func=lambda x: aprendices_lista[x])

    # Función para reemplazar marcadores en las tablas del Word
    def reemplazar_en_tabla(doc, reemplazos):
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for k, v in reemplazos.items():
                        if k in cell.text:
                            cell.text = cell.text.replace(k, str(v))

    if st.button("🚀 Generar Plan Concertado"):
        # Obtener los datos del aprendiz seleccionado
        row = df.iloc[aprendiz_seleccionado_idx]
        
        # Cargar documento Word plantilla
        doc = Document(word_template_file)
        
        # Diccionario de reemplazos básicos
        reemplazos = {
            "«Nombre»": row[col_nombre],
            "«Apellidos»": row[col_apellido],
            "«Tipo_de_Doc»": row[col_tipo_doc],
            "«N_Documento»": row[col_doc],
            "JHON CUENTAS DE CARO": instructor,  # O marcador si está en la plantilla
        }
        
        # Lógica para la sección de entregas (SI / NO) en el reporte final
        for i in range(1, 11):
            tag_si = f"«Act_{i}_Si»"
            tag_no = f"«Act_{i}_No»"
            
            if tipo_documento.startswith("Inicial"):
                # En el plan inicial se dejan en blanco
                reemplazos[tag_si] = ""
                reemplazos[tag_no] = ""
            else:
                # En el plan final se marca (aquí puedes ajustar la condición de entrega)
                reemplazos[tag_si] = "X"
                reemplazos[tag_no] = ""

        # Aplicar reemplazos
        reemplazar_en_tabla(doc, reemplazos)
        
        # Guardar archivo generado en un buffer de memoria
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
