import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄")

st.title("📄 Generador de Planes de Trabajo - SENA")
st.write("Sube el archivo de **Reporte de Juicios Evaluativos (.xls/.xlsx)** para generar automáticamente los documentos Word de los **aprendices en formación**.")

# Subir archivo Excel desde la interfaz web
uploaded_excel = st.file_uploader("Cargar Reporte de Juicios Evaluativos (Excel)", type=["xls", "xlsx"])

if uploaded_excel is not None:
    try:
        # Leer el Excel omitiendo los encabezados de metadatos (fila 13 en adelante)
        df_raw = pd.read_excel(uploaded_excel, skiprows=12)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        # 1. Filtrar únicamente a los aprendices con estado "EN FORMACION"
        df_filtrado = df_raw[df_raw['Estado'].str.upper() == 'EN FORMACION'].copy()
        
        # 2. Obtener la lista de aprendices únicos en formación
        aprendices = df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Estado']].drop_duplicates()
        st.success(f"✅ Se encontraron **{len(aprendices)}** aprendices con estado **EN FORMACION**.")
        
        # Mostrar vista previa de los aprendices que se van a procesar
        st.dataframe(aprendices[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos']], use_container_width=True)
        
        if st.button("🚀 Generar Planes de Trabajo en ZIP"):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for idx, aprendiz in aprendices.iterrows():
                    num_doc = str(aprendiz['Número de Documento'])
                    nombre_completo = f"{aprendiz['Nombre']} {aprendiz['Apellidos']}"
                    
                    # Cargar la plantilla Word incluida en el repositorio
                    doc = Document("Plan de trabajo .docx")
                    
                    # Mapear los datos generales del aprendiz
                    reemplazos = {
                        "«Nombre»": str(aprendiz['Nombre']),
                        "«Apellidos»": str(aprendiz['Apellidos']),
                        "«Tipo_de_Doc»": str(aprendiz['Tipo de Documento']),
                        "«N_Documento»": num_doc,
                    }
                    
                    # Filtrar juicios evaluativos de este aprendiz específico
                    df_aprendiz = df_filtrado[df_filtrado['Número de Documento'] == aprendiz['Número de Documento']].reset_index(drop=True)
                    
                    # Asignar los juicios a los marcadores Act_1_Si, Act_1_No, etc.
                    for i in range(1, 11):
                        si_key = f"«Act_{i}_Si»"
                        no_key = f"«Act_{i}_No»"
                        
                        if i - 1 < len(df_aprendiz):
                            juicio = str(df_aprendiz.loc[i - 1, 'Juicio de Evaluación']).upper()
                            if "APROBADO" in juicio:
                                reemplazos[si_key] = "X"
                                reemplazos[no_key] = ""
                            else:
                                reemplazos[si_key] = ""
                                reemplazos[no_key] = "X"
                        else:
                            reemplazos[si_key] = ""
                            reemplazos[no_key] = ""

                    # Reemplazar valores en todos los párrafos del documento Word
                    for p in doc.paragraphs:
                        for k, v in reemplazos.items():
                            if k in p.text:
                                p.text = p.text.replace(k, v)
                                
                    # Reemplazar valores en las tablas del documento Word
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for k, v in reemplazos.items():
                                    if k in cell.text:
                                        cell.text = cell.text.replace(k, v)

                    # Guardar el documento del aprendiz en memoria interna
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    
                    filename = f"Plan_Trabajo_{num_doc}_{aprendiz['Nombre']}_{aprendiz['Apellidos']}.docx"
                    zip_file.writestr(filename, doc_io.getvalue())
            
            # Botón de descarga para el archivo ZIP resultante
            st.download_button(
                label="📦 Descargar Documentos de Aprendices en Formación (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Planes_Trabajo_Aprendices_En_Formacion.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
