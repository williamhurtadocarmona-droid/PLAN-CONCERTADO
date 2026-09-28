import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄", layout="wide")

st.title("📄 Generador de Planes de Trabajo - SENA")
st.write("Sube el archivo de **Reporte de Juicios Evaluativos (.xls/.xlsx)**, selecciona los **Resultados de Aprendizaje** a evaluar y genera los documentos en Word.")

# Subir archivo Excel desde la interfaz web
uploaded_excel = st.file_uploader("Cargar Reporte de Juicios Evaluativos (Excel)", type=["xls", "xlsx"])

if uploaded_excel is not None:
    try:
        # 1. Extraer metadatos exactos de las primeras 12 filas
        df_meta = pd.read_excel(uploaded_excel, header=None, nrows=12)
        
        denominacion_programa = ""
        numero_ficha = ""
        
        for idx, row in df_meta.iterrows():
            label = str(row[0]).strip()
            if label == "Ficha de Caracterización:":
                numero_ficha = str(row[2]).strip() if pd.notna(row[2]) else str(row[1]).strip()
            elif label == "Denominación:":
                denominacion_programa = str(row[2]).strip() if pd.notna(row[2]) else str(row[1]).strip()

        st.info(f"📌 **Programa:** {denominacion_programa} | **Ficha de Caracterización:** {numero_ficha}")

        # 2. Leer la tabla de juicios evaluativos (a partir de la fila 13)
        df_raw = pd.read_excel(uploaded_excel, skiprows=12)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        # 3. Filtrar únicamente a los aprendices con estado "EN FORMACION"
        df_filtrado = df_raw[df_raw['Estado'].str.upper() == 'EN FORMACION'].copy()
        
        # 4. Obtener lista de Resultados de Aprendizaje (RAPs) disponibles
        raps_disponibles = sorted(df_filtrado['Resultado de Aprendizaje'].dropna().unique().tolist())
        
        st.subheader("🎯 Selección de Resultados de Aprendizaje (RAP)")
        raps_seleccionados = st.multiselect(
            "Selecciona uno o más Resultados de Aprendizaje que deseas incluir en los planes de trabajo:",
            options=raps_disponibles,
            default=raps_disponibles[:3] if len(raps_disponibles) >= 3 else raps_disponibles,
            help="Puedes seleccionar hasta 10 RAPs para que coincidan con la estructura de la plantilla."
        )

        if len(raps_seleccionados) == 0:
            st.warning("⚠️ Debes seleccionar al menos un Resultado de Aprendizaje para continuar.")
        elif len(raps_seleccionados) > 10:
            st.error("❌ Has seleccionado más de 10 RAPs. La plantilla actual admite máximo 10 actividades por documento.")
        else:
            # 5. Obtener la lista de aprendices únicos en formación
            aprendices = df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Estado']].drop_duplicates()
            st.success(f"✅ Aprendices a procesar: **{len(aprendices)}** | RAPs seleccionados: **{len(raps_seleccionados)}**")
            
            with st.expander("👁️ Ver lista de aprendices a procesar"):
                st.dataframe(aprendices[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos']], use_container_width=True)

            if st.button("🚀 Generar Planes de Trabajo en ZIP"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for idx, aprendiz in aprendices.iterrows():
                        num_doc = str(aprendiz['Número de Documento'])
                        
                        # Cargar la plantilla Word incluida en el repositorio
                        doc = Document("Plan de trabajo .docx")
                        
                        # Diccionario de reemplazos generales
                        reemplazos = {
                            "«Nombre»": str(aprendiz['Nombre']),
                            "«Apellidos»": str(aprendiz['Apellidos']),
                            "«Tipo_de_Doc»": str(aprendiz['Tipo de Documento']),
                            "«N_Documento»": num_doc,
                            
                            # Programa y Ficha
                            "TECNICO INSTALACION SISTEMAS ELECTRICOS RESIDENCIALES Y COMERCIALES": denominacion_programa,
                            "«Programa»": denominacion_programa,
                            "«Denominacion»": denominacion_programa,
                            "837101": numero_ficha,
                            "«Ficha»": numero_ficha,
                            "«Numero_Ficha»": numero_ficha,
                        }
                        
                        # Filtrar juicios evaluativos del aprendiz actual
                        df_aprendiz = df_filtrado[df_filtrado['Número de Documento'] == aprendiz['Número de Documento']]
                        
                        # Procesar cada uno de los RAPs seleccionados (hasta 10)
                        for i in range(1, 11):
                            rap_key = f"«Act_{i}»"          # Si agregas marcador del texto del RAP
                            si_key = f"«Act_{i}_Si»"
                            no_key = f"«Act_{i}_No»"
                            
                            if i - 1 < len(raps_seleccionados):
                                rap_actual = raps_seleccionados[i - 1]
                                reemplazos[rap_key] = rap_actual
                                
                                # Buscar el juicio del aprendiz para este RAP
                                fila_rap = df_aprendiz[df_aprendiz['Resultado de Aprendizaje'] == rap_actual]
                                
                                if not fila_rap.empty:
                                    juicio = str(fila_rap.iloc[0]['Juicio de Evaluación']).upper()
                                    if "APROBADO" in juicio:
                                        reemplazos[si_key] = "X"
                                        reemplazos[no_key] = ""
                                    else:
                                        reemplazos[si_key] = ""
                                        reemplazos[no_key] = "X"
                                else:
                                    # Si el RAP no está registrado para el aprendiz
                                    reemplazos[si_key] = ""
                                    reemplazos[no_key] = "X"
                            else:
                                # Dejar vacíos las filas sobrantes
                                reemplazos[rap_key] = ""
                                reemplazos[si_key] = ""
                                reemplazos[no_key] = ""

                        # Reemplazar en párrafos
                        for p in doc.paragraphs:
                            for k, v in reemplazos.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, v)
                                    
                        # Reemplazar en tablas
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
                
                # Botón para descargar el ZIP resultante
                st.download_button(
                    label="📦 Descargar Documentos de Aprendices (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Planes_Trabajo_Ficha_{numero_ficha}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
