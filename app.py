import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄", layout="wide")

st.title("📄 Generador de Planes de Trabajo - SENA")
st.write("Sube el archivo de **Reporte de Juicios Evaluativos (.xls/.xlsx)**, configura los **Resultados de Aprendizaje**, **Actividades**, **Forma de Entrega** y la **Etapa del Plan**.")

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
        
        st.subheader("🎯 1. Selección de Resultados de Aprendizaje (RAP)")
        raps_seleccionados = st.multiselect(
            "Selecciona los Resultados de Aprendizaje que deseas incluir en el Plan de Trabajo:",
            options=raps_disponibles,
            default=raps_disponibles[:3] if len(raps_disponibles) >= 3 else raps_disponibles,
            help="Puedes seleccionar hasta 10 RAPs."
        )

        num_raps = len(raps_seleccionados)

        if num_raps == 0:
            st.warning("⚠️ Debes seleccionar al menos un Resultado de Aprendizaje para continuar.")
        elif num_raps > 10:
            st.error("❌ Has seleccionado más de 10 RAPs. La plantilla actual admite un máximo de 10 actividades.")
        else:
            # 5. Opciones del Tipo de Plan (Inicial vs Final)
            st.subheader("📋 2. Estado del Plan de Trabajo")
            tipo_plan = st.radio(
                "Selecciona el momento de generación del Plan de Trabajo:",
                options=["Plan Inicial", "Plan Final"],
                index=0,
                horizontal=True,
                help="En 'Plan Inicial' los estados se completan automáticamente según el Excel. En 'Plan Final' puedes definir la entrega manual de cada actividad."
            )

            # 6. Configurar Actividades a desarrollar, Forma de Entrega y Estado de Entrega (si es Final)
            st.subheader("📝 3. Configurar Actividades y Entregas")
            
            actividades_por_rap = {}
            entrega_por_rap = {}
            estado_entrega_final = {}

            for i, rap in enumerate(raps_seleccionados, 1):
                st.markdown("---")
                st.markdown(f"**Actividad {i}:** `{rap}`")
                
                # Ajustar columnas dinámicamente según si es Plan Inicial o Final
                if tipo_plan == "Plan Final":
                    col1, col2, col3 = st.columns([3, 1, 1])
                else:
                    col1, col2 = st.columns([3, 1])
                
                with col1:
                    actividades_por_rap[rap] = st.text_area(
                        f"Descripción de la Actividad {i}",
                        value=f"Desarrollar guía de aprendizaje y evidencias prácticas de: {rap.split('-')[-1].strip()}",
                        key=f"act_rap_{i}",
                        height=80
                    )
                
                with col2:
                    entrega_por_rap[rap] = st.radio(
                        f"Forma de Entrega {i}",
                        options=["Física", "Digital"],
                        index=1,
                        key=f"entrega_rap_{i}"
                    )
                
                if tipo_plan == "Plan Final":
                    with col3:
                        estado_entrega_final[rap] = st.radio(
                            f"¿Entregó Actividad {i}?",
                            options=["SÍ", "NO"],
                            index=0,
                            key=f"estado_entrega_{i}"
                        )

            # 7. Obtener la lista de aprendices únicos en formación
            aprendices = df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Estado']].drop_duplicates()
            st.success(f"✅ Aprendices a procesar: **{len(aprendices)}** | Tipo de Plan: **{tipo_plan}** | RAPs a evaluar: **{num_raps}**")
            
            with st.expander("👁️ Ver lista de aprendices a procesar"):
                st.dataframe(aprendices[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos']], use_container_width=True)

            if st.button("🚀 Generar Planes de Trabajo en ZIP"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for idx, aprendiz in aprendices.iterrows():
                        num_doc = str(aprendiz['Número de Documento'])
                        
                        # Cargar la plantilla Word incluida en el repositorio
                        doc = Document("Plan de trabajo .docx")
                        
                        # Diccionario de reemplazos generales del encabezado
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
                        
                        # Reemplazar encabezados en párrafos
                        for p in doc.paragraphs:
                            for k, v in reemplazos.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, v)
                                    
                        # Reemplazar encabezados en la primera tabla
                        if len(doc.tables) > 0:
                            for row in doc.tables[0].rows:
                                for cell in row.cells:
                                    for k, v in reemplazos.items():
                                        if k in cell.text:
                                            cell.text = cell.text.replace(k, v)

                        # Filtrar juicios evaluativos del aprendiz actual
                        df_aprendiz = df_filtrado[df_filtrado['Número de Documento'] == aprendiz['Número de Documento']]
                        
                        # Modificar la segunda tabla (Tabla 1: Descriptores de la Ruta de Aprendizaje)
                        if len(doc.tables) > 1:
                            tabla_actividades = doc.tables[1]
                            
                            # Rellenar filas de actividades
                            for i in range(num_raps):
                                row_idx = i + 3
                                if row_idx < len(tabla_actividades.rows):
                                    row_cells = tabla_actividades.rows[row_idx].cells
                                    rap_actual = raps_seleccionados[i]
                                    
                                    # Columna 0: Resultados de Aprendizaje
                                    row_cells[0].text = rap_actual
                                    
                                    # Columna 1: No Actividad
                                    row_cells[1].text = str(i + 1)
                                    
                                    # Columna 2: Actividades a desarrollar
                                    row_cells[2].text = actividades_por_rap.get(rap_actual, "")
                                    
                                    # Forma de Entrega: Columna 3 (Física) / Columna 4 (Digital)
                                    tipo_entrega = entrega_por_rap.get(rap_actual, "Digital")
                                    if tipo_entrega == "Física":
                                        row_cells[3].text = "X"
                                        row_cells[4].text = ""
                                    else:
                                        row_cells[3].text = ""
                                        row_cells[4].text = "X"

                                    # Determinación de entrega según si es Plan Inicial o Plan Final
                                    if tipo_plan == "Plan Final":
                                        entrego = estado_entrega_final.get(rap_actual, "SÍ")
                                        if entrego == "SÍ":
                                            row_cells[7].text = "X"  # Columna 7: SI
                                            row_cells[8].text = ""   # Columna 8: NO
                                        else:
                                            row_cells[7].text = ""   # Columna 7: SI
                                            row_cells[8].text = "X"  # Columna 8: NO
                                    else:
                                        # Plan Inicial: Lee el reporte del Excel
                                        fila_rap = df_aprendiz[df_aprendiz['Resultado de Aprendizaje'] == rap_actual]
                                        if not fila_rap.empty:
                                            juicio = str(fila_rap.iloc[0]['Juicio de Evaluación']).upper()
                                            if "APROBADO" in juicio:
                                                row_cells[7].text = "X"  # Columna 7: SI
                                                row_cells[8].text = ""   # Columna 8: NO
                                            else:
                                                row_cells[7].text = ""   # Columna 7: SI
                                                row_cells[8].text = "X"  # Columna 8: NO
                                        else:
                                            row_cells[7].text = ""
                                            row_cells[8].text = "X"

                            # ELIMINAR LAS FILAS SOBRANTES
                            for row_to_remove_idx in range(12, num_raps + 2, -1):
                                if row_to_remove_idx < len(tabla_actividades.rows):
                                    tr = tabla_actividades.rows[row_to_remove_idx]._tr
                                    tr.getparent().remove(tr)

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
                    file_name=f"Planes_Trabajo_{tipo_plan.replace(' ', '_')}_Ficha_{numero_ficha}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
