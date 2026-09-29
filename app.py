import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import re
import os
import tempfile
import subprocess

# PyPDF para la unificación de páginas
from pypdf import PdfWriter

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄", layout="wide")

st.title("📄 Generador de Planes de Trabajo - SENA")
st.write("Sube el archivo de **Reporte de Juicios Evaluativos (.xls/.xlsx)**, selecciona el **Programa**, configura las **Actividades por RAP**, **Forma de Entrega** y el **Estado del Plan**.")

# Subir archivo Excel desde la interfaz web
uploaded_excel = st.file_uploader("Cargar Reporte de Juicios Evaluativos (Excel)", type=["xls", "xlsx"])

def convertir_docx_fiel_a_pdf(doc_bytes_io):
    """
    Convierte el archivo Word (.docx) poblado directamente a PDF para preservar 
    el logo oficial SENA, colores, bordes y tipografía exacta.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "documento_sena.docx")
        pdf_path = os.path.join(tmpdir, "documento_sena.pdf")
        
        with open(docx_path, "wb") as f:
            f.write(doc_bytes_io.getvalue())

        # 1. Intentar con LibreOffice (Sistemas Linux / Servidor / Streamlit Cloud)
        try:
            cmd = f"libreoffice --headless --convert-to pdf {docx_path} --outdir {tmpdir}"
            subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return io.BytesIO(f.read())
        except Exception:
            pass

        # 2. Intentar con soffice
        try:
            cmd = f"soffice --headless --convert-to pdf {docx_path} --outdir {tmpdir}"
            subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return io.BytesIO(f.read())
        except Exception:
            pass

        # 3. Intentar con docx2pdf (Entornos Windows con MS Word)
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return io.BytesIO(f.read())
        except Exception:
            pass

    return None

if uploaded_excel is not None:
    try:
        # 1. Menú desplegable para selección del Programa (TOTF vs TMMI)
        st.subheader("📚 Selección del Programa de Formación")
        programa_seleccionado = st.selectbox(
            "Selecciona la Especialidad/Programa:",
            options=["TOTF - Operación en Torno y Fresadora", "TMMI - Mantenimiento Mecánico Industrial"],
            index=0,
            help="Selecciona el programa correspondiente."
        )

        # 2. Extraer metadatos exactos del Excel (filas 0 a 11)
        df_meta = pd.read_excel(uploaded_excel, header=None, nrows=12)
        
        numero_ficha = ""
        for idx, row in df_meta.iterrows():
            label = str(row[0]).strip()
            if label == "Ficha de Caracterización:":
                numero_ficha = str(row[2]).strip() if pd.notna(row[2]) else str(row[1]).strip()

        # Configuración automática de Programa y Proyecto según la selección (TOTF o TMMI)
        if "TOTF" in programa_seleccionado:
            denominacion_programa = "OPERACION EN TORNO Y FRESADORA"
            proyecto_default = "OPTIMIZACIÓN EN LA FABRICACIÓN DE COMPONENTES MECÁNICOS EN TORNO Y FRESADORA EN LAS INDUSTRIAS DEL ATLÁNTICO"
        else:
            denominacion_programa = "MECANICA DE MAQUINARIA INDUSTRIAL"
            proyecto_default = "IMPLEMENTACIÓN DEL PROGRAMA DE MANTENIMIENTO MECÁNICO INDUSTRIAL EN INDUSTRIAS Y CENTROS DE FORMACION SENA REGIONAL ATLÁNTICO."

        st.info(f"📌 **Especialidad:** {programa_seleccionado} | **Programa de Formación:** {denominacion_programa} | **Ficha:** {numero_ficha}")

        # 3. Leer la tabla de juicios evaluativos (a partir de la fila 13)
        df_raw = pd.read_excel(uploaded_excel, skiprows=12)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        # 4. Filtrar únicamente a los aprendices con estado "EN FORMACION"
        df_filtrado = df_raw[df_raw['Estado'].str.upper() == 'EN FORMACION'].copy()
        
        # 5. Información General del Proyecto Formativo y Fase
        st.subheader("🛠️ 1. Datos del Proyecto Formativo")
        col_proj1, col_proj2 = st.columns(2)
        
        with col_proj1:
            proyecto_formativo_input = st.text_input(
                "Proyecto Formativo:",
                value=proyecto_default,
                placeholder="Escribe el nombre del Proyecto Formativo...",
                help="Se autocompleta automáticamente según la opción seleccionada (TOTF o TMMI)."
            )
            
        with col_proj2:
            fase_proyecto_input = st.selectbox(
                "Fase del Proyecto:",
                options=["Análisis", "Planeación", "Ejecución", "Evaluación"],
                index=2, # Ejecución por defecto
                help="Selecciona la Fase del Proyecto correspondiente."
            )

        # 6. Selección de Competencia y Resultados de Aprendizaje (RAP)
        st.subheader("🎯 2. Selección de Competencia y Resultados de Aprendizaje (RAP)")
        
        # Actividades para TOTF RAP 1 (TORNO CONVENCIONAL)
        actividades_totf_rap1 = [
            "Preparación de herramientas y equipo: Identificar el tipo de herramienta de corte según material (acero, aluminio, bronce, etc.) y tipo de operación (cilindrado, refrentado, ranurado, roscado). Acondicionar y afilar herramientas de corte aplicando ángulos de incidencia, desprendimiento y radio punta adecuados. Seleccionar y preparar los elementos de sujeción (mordazas, portaherramientas, contrapunto).",
            "Elaboración de la orden operacional: Analizar un plano técnico de la pieza: dimensiones, tolerancias, acabados y ajustes. Secuenciar operaciones de mecanizado en una orden de trabajo (operaciones preliminares → desbaste → semiacabado → acabado). Calcular velocidades y avances adecuados de acuerdo al material.",
            "Puesta a punto del torno convencional: Montar la pieza en el torno, verificando sujeción y alineación. Instalar correctamente la herramienta, calibrar alturas y cotas iniciales. Realizar pruebas de giro y verificación de seguridad.",
            "Ejecución de operaciones de torneado: Cilindrado, refrentado, taladrado, roscado y tronzado según plano técnico y orden operacional."
        ]

        # Actividades para TOTF RAP 2 (FRESADORA CONVENCIONAL)
        actividades_totf_rap2 = [
            "Preparación y puesta a punto de máquina y herramientas: Seleccionar cortadores (fresas cilíndricas, de disco, de punta esférica, de ranura) según material y operación. Montar la pieza correctamente con sistemas de sujeción (mordazas, bridas, divisores). Calibrar recorridos y verificar movimientos de avance y corte.",
            "Ejecución del fresado según orden operacional: Interpretar plano y definir secuencia de operaciones de fresado (planeado, ranurado, engranaje, roscado). Ejecutar operaciones siguiendo el plan de trabajo.",
            "Control dimensional y verificación: Medir cotas de la pieza fresada con instrumentos de metrología. Validar tolerancias, geometrías y superficie de acuerdo al plano.",
            "Identificación de fallas y mejoras: Detectar defectos como vibraciones, desviaciones de medida, mala sujeción o acabados deficientes. Reportar fallas en un formato técnico y proponer acciones correctivas.",
            "Seguridad y protección en fresado: Usar adecuadamente equipos de protección personal (monogafas, guantes, protección auditiva). Garantizar buenas prácticas en la manipulación de virutas y refrigerante."
        ]

        # Actividades para TMMI (RAP 3)
        actividades_tmmi_rap3 = [
            "Fabricar elementos mecánicos aplicando procesos de mecanizado con torno convencional."
        ]

        if "TOTF" in programa_seleccionado:
            competencia_totf = st.selectbox(
                "Selecciona la Competencia:",
                options=["290201211 - MECANIZAR PIEZA INDUSTRIAL DE ACUERDO CON TÉCNICAS MANUALES Y SEMIAUTOMÁTICAS"],
                index=0
            )
            
            rap_totf_seleccionado = st.selectbox(
                "Selecciona el Resultado de Aprendizaje (RAP):",
                options=[
                    "694494 - 1.OPERAR TORNO CONVENCIONAL DE ACUERDO CON PROCEDIMIENTOS TÉCNICOS Y NORMATIVA.",
                    "694495 - 2.OPERAR FRESADORA CONVENCIONAL DE ACUERDO CON PROCEDIMIENTOS TÉCNICOS Y NORMATIVA"
                ],
                index=0
            )
            
            raps_seleccionados = [rap_totf_seleccionado]
            
            if "694494" in rap_totf_seleccionado or "TORNO" in rap_totf_seleccionado.upper():
                lista_actividades_base = actividades_totf_rap1
            else:
                lista_actividades_base = actividades_totf_rap2
        else:
            competencia_tmmi = st.selectbox(
                "Selecciona la Competencia:",
                options=["REPARAR EQUIPOS SEGUN PROCEDIMIENTOS Y MANUALES TECNICOS."],
                index=0
            )
            
            rap_tmmi_seleccionado = st.selectbox(
                "Selecciona el Resultado de Aprendizaje (RAP):",
                options=["OPERAR MÁQUINAS Y HERRAMIENTAS CONVENCIONALES SEGÚN ESPECIFICACIONES TÉCNICAS."],
                index=0
            )
            
            raps_seleccionados = [rap_tmmi_seleccionado]
            lista_actividades_base = actividades_tmmi_rap3

        # 7. Opciones del Tipo de Plan (Inicial vs Final)
        st.subheader("📋 3. Estado del Plan de Trabajo")
        tipo_plan = st.radio(
            "Selecciona el momento de generación del Plan de Trabajo:",
            options=["Plan Inicial", "Plan Final"],
            index=0,
            horizontal=True
        )

        # 8. Opciones de Configuración Masiva
        st.subheader("⚡ 4. Aplicación Masiva (Opcional)")
        col_m1, col_m2 = st.columns(2)
        
        with col_m1:
            entrega_masiva = st.selectbox(
                "Forma de Entrega para TODAS las actividades:",
                options=["Sin cambio masivo (Personalizar abajo)", "Física (Todas)", "Digital (Todas)"],
                index=0
            )
        
        with col_m2:
            if tipo_plan == "Plan Final":
                estado_masivo = st.selectbox(
                    "¿Entregó la actividad? para TODAS (Plan Final):",
                    options=["Sin cambio masivo (Personalizar abajo)", "SÍ (Todos aprobaron / entregaron)", "NO (Ninguno entregó)"],
                    index=0
                )
            else:
                estado_masivo = "Sin cambio masivo"

        # 9. Configurar Actividades en celdas independientes
        st.subheader("📝 5. Configurar Actividades e Individualizar")
        
        actividades_desc = []
        entrega_por_act = []
        estado_entrega_final = []

        idx_entrega_default = 0 if "Física" in entrega_masiva else 1
        idx_estado_default = 0 if "SÍ" in estado_masivo else 1

        for i, act_txt in enumerate(lista_actividades_base, 1):
            st.markdown("---")
            st.markdown(f"**Actividad {i}:**")
            
            if tipo_plan == "Plan Final":
                col1, col2, col3 = st.columns([3, 1, 1])
            else:
                col1, col2 = st.columns([3, 1])
            
            with col1:
                val_act = st.text_area(
                    f"Descripción de la Celda {i}",
                    value=act_txt,
                    key=f"act_celda_{i}",
                    height=90
                )
                actividades_desc.append(val_act)
            
            with col2:
                sub_idx_entrega = idx_entrega_default if "Todas" in entrega_masiva else 1
                ent_val = st.radio(
                    f"Forma de Entrega Celda {i}",
                    options=["Física", "Digital"],
                    index=sub_idx_entrega,
                    key=f"entrega_celda_{i}"
                )
                entrega_por_act.append(ent_val)
            
            if tipo_plan == "Plan Final":
                with col3:
                    sub_idx_estado = idx_estado_default if ("Todos" in estado_masivo or "Ninguno" in estado_masivo) else 0
                    est_val = st.radio(
                        f"¿Entregó Celda {i}?",
                        options=["SÍ", "NO"],
                        index=sub_idx_estado,
                        key=f"estado_celda_{i}"
                    )
                    estado_entrega_final.append(est_val)

        # 10. Obtener la lista de aprendices únicos en formación
        aprendices = df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Estado']].drop_duplicates()
        st.success(f"✅ Aprendices a procesar: **{len(aprendices)}** | Programa: **{programa_seleccionado.split(' - ')[0]}** | Actividades a evaluar: **{len(actividades_desc)}**")
        
        with st.expander("👁️ Ver lista de aprendices a procesar"):
            st.dataframe(aprendices[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos']], use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Generación y Descarga de Documentos")
        
        col_btn1, col_btn2 = st.columns(2)

        def generar_doc_original_poblado(aprendiz):
            num_doc = str(aprendiz['Número de Documento'])
            
            # Carga la plantilla oficial en Word respetando logo SENA, colores e imágenes
            doc = Document("Plan de trabajo .docx")
            
            reemplazos = {
                "«Nombre»": str(aprendiz['Nombre']),
                "«Apellidos»": str(aprendiz['Apellidos']),
                "«Tipo_de_Doc»": str(aprendiz['Tipo de Documento']),
                "«N_Documento»": num_doc,
                "TECNICO INSTALACION SISTEMAS ELECTRICOS RESIDENCIALES Y COMERCIALES": denominacion_programa,
                "«Programa»": denominacion_programa,
                "«Denominacion»": denominacion_programa,
                "837101": numero_ficha,
                "«Ficha»": numero_ficha,
                "«Numero_Ficha»": numero_ficha,
                "«Proyecto_Formativo»": proyecto_formativo_input,
                "«Proyecto»": proyecto_formativo_input,
                "«Fase_Proyecto»": fase_proyecto_input,
                "«Fase»": fase_proyecto_input,
            }
            
            for p in doc.paragraphs:
                for k, v in reemplazos.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        
            if len(doc.tables) > 0:
                t0 = doc.tables[0]
                for row in t0.rows:
                    for cell in row.cells:
                        for k, v in reemplazos.items():
                            if k in cell.text:
                                cell.text = cell.text.replace(k, v)
                
                if len(t0.rows) > 2 and len(t0.rows[2].cells) > 5:
                    if "Proyecto Formativo:" in t0.rows[2].cells[5].text:
                        t0.rows[2].cells[6].text = proyecto_formativo_input
                        
                if len(t0.rows) > 3 and len(t0.rows[3].cells) > 1:
                    if "Fase del" in t0.rows[3].cells[0].text or "Fase" in t0.rows[3].cells[0].text:
                        t0.rows[3].cells[1].text = fase_proyecto_input

            if len(doc.tables) > 1:
                tabla_actividades = doc.tables[1]
                rap_actual_str = raps_seleccionados[0] if len(raps_seleccionados) > 0 else ""
                
                for i, act_descripcion in enumerate(actividades_desc):
                    row_idx = i + 3
                    if row_idx < len(tabla_actividades.rows):
                        row_cells = tabla_actividades.rows[row_idx].cells
                        row_cells[0].text = rap_actual_str
                        row_cells[1].text = str(i + 1)
                        row_cells[2].text = act_descripcion
                        
                        tipo_entrega = entrega_por_act[i]
                        if tipo_entrega == "Física":
                            row_cells[3].text = "X"
                            row_cells[4].text = ""
                        else:
                            row_cells[3].text = ""
                            row_cells[4].text = "X"

                        if tipo_plan == "Plan Final":
                            entrego = estado_entrega_final[i]
                            if entrego == "SÍ":
                                row_cells[7].text = "X"
                                row_cells[8].text = ""
                            else:
                                row_cells[7].text = ""
                                row_cells[8].text = "X"
                        else:
                            row_cells[7].text = ""
                            row_cells[8].text = ""

                total_filas_insertadas = len(actividades_desc)
                for row_to_remove_idx in range(12, total_filas_insertadas + 2, -1):
                    if row_to_remove_idx < len(tabla_actividades.rows):
                        tr = tabla_actividades.rows[row_to_remove_idx]._tr
                        tr.getparent().remove(tr)

            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            return doc_io

        # BOTÓN 1: Generar PDF Unificado desde la plantilla original
        with col_btn1:
            if st.button("📄 Generar UN SOLO ARCHIVO PDF Unificado", use_container_width=True):
                pdf_merger = PdfWriter()
                converted_count = 0

                for idx, aprendiz in aprendices.iterrows():
                    doc_bytes = generar_doc_original_poblado(aprendiz)
                    pdf_io = convertir_docx_fiel_a_pdf(doc_bytes)
                    
                    if pdf_io:
                        pdf_merger.append(pdf_io)
                        converted_count += 1

                if converted_count > 0:
                    final_pdf_buffer = io.BytesIO()
                    pdf_merger.write(final_pdf_buffer)
                    pdf_merger.close()
                    final_pdf_buffer.seek(0)

                    sigla_prog = programa_seleccionado.split(' - ')[0]
                    st.download_button(
                        label="⬇️ Descargar PDF Unificado Completo (Original SENA)",
                        data=final_pdf_buffer.getvalue(),
                        file_name=f"Planes_Trabajo_UNIFICADO_{sigla_prog}_{tipo_plan.replace(' ', '_')}_Ficha_{numero_ficha}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error("⚠️ Para convertir los documentos de Word a PDF manteniendo exactamente el logo SENA y el formato institucional original, instala LibreOffice en tu servidor ejecutando: `apt-get install -y libreoffice` (en Linux/Docker) o instala Microsoft Word en Windows.")

        # BOTÓN 2: Generar Planes en ZIP (.docx con plantilla original completa)
        with col_btn2:
            if st.button("📦 Generar Planes en ZIP (.docx indv.)", use_container_width=True):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for idx, aprendiz in aprendices.iterrows():
                        num_doc = str(aprendiz['Número de Documento'])
                        doc_bytes = generar_doc_original_poblado(aprendiz)
                        
                        sigla_prog = programa_seleccionado.split(' - ')[0]
                        filename = f"Plan_Trabajo_{sigla_prog}_{num_doc}_{aprendiz['Nombre']}_{aprendiz['Apellidos']}.docx"
                        zip_file.writestr(filename, doc_bytes.getvalue())
                
                sigla_prog = programa_seleccionado.split(' - ')[0]
                st.download_button(
                    label="⬇️ Descargar ZIP de Documentos Word (Plantilla SENA Original)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Planes_Trabajo_{sigla_prog}_{tipo_plan.replace(' ', '_')}_Ficha_{numero_ficha}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
