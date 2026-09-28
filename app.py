import streamlit as st
import pandas as pd
from docx import Document
import io
import os
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Generador Plan Concertado - SENA PDF", layout="wide")

st.title("📋 Generador de Planes Concertados (PDF Unificado)")
st.write("Genera los planes concertados individuales de cada aprendiz y consolidalos en un solo archivo PDF.")

PLANTILLA_PATH = "Plan de trabajo .docx"

# 1. Cargar archivo Excel
excel_file = st.file_uploader("1. Cargar archivo Excel de Juicios Evaluativos", type=["xlsx", "xls"])

if excel_file:
    try:
        # Fila 13 (index 12) contiene los títulos del reporte
        df = pd.read_excel(excel_file, header=12)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

    df.columns = [str(col).strip() for col in df.columns]
    cols_lista = list(df.columns)

    # Buscar índices automáticos de columnas
    def buscar_idx(patrones):
        for pat in patrones:
            for idx, col in enumerate(cols_lista):
                if pat.lower() in str(col).lower():
                    return idx
        return 0

    idx_nom = buscar_idx(["nombre"])
    idx_ape = buscar_idx(["apellido"])
    idx_tipo = buscar_idx(["tipo"])
    idx_doc = buscar_idx(["numero", "documento", "identifica"])
    idx_est = buscar_idx(["estado"])
    idx_comp = buscar_idx(["competencia"])
    idx_rap = buscar_idx(["resultado", "rap"])

    st.subheader("📝 2. Configuración General")
    col1, col2, col3 = st.columns(3)

    with col1:
        instructor = st.text_input("Instructor:", value="WILLIAM SANTIAGO HURTADO CARMONA")
        programa = st.text_input("Programa de Formación:", value="TECNICO INSTALACION SISTEMAS ELECTRICOS RESIDENCIALES Y COMERCIALES")
    
    with col2:
        proyecto = st.text_input("Proyecto Formativo:", value="")
        fase_proyecto = st.text_input("Fase del Proyecto:", value="")

    with col3:
        forma_entrega = st.selectbox("Forma de Entrega:", ["Digital", "Física"])
        tipo_reporte = st.radio("Tipo de Reporte:", ["Inicial (En Blanco)", "Final (Con SI)"])

    st.subheader("🎯 3. Selección de Competencia y Resultado de Aprendizaje")
    col_c1, col_c2 = st.columns(2)

    with col_c1:
        col_competencia = st.selectbox("Columna Competencia:", cols_lista, index=idx_comp)
        comps_unicas = df[col_competencia].dropna().unique().tolist()
        competencia_sel = st.selectbox("Seleccione la Competencia:", comps_unicas) if comps_unicas else ""

    with col_c2:
        col_rap = st.selectbox("Columna Resultado de Aprendizaje:", cols_lista, index=idx_rap)
        raps_unicos = df[col_rap].dropna().unique().tolist()
        rap_sel = st.selectbox("Seleccione el Resultado de Aprendizaje:", raps_unicos) if raps_unicos else ""

    # Filtrar por Aprendices "EN FORMACION"
    col_est = cols_lista[idx_est]
    df_formacion = df[df[col_est].astype(str).str.upper().str.contains("FORMACI", na=False)].copy()
    df_aprendices = df_formacion.drop_duplicates(subset=[cols_lista[idx_doc]]).reset_index(drop=True)

    st.subheader("👥 4. Aprendices en Formación Detectados")
    
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        col_nom = st.selectbox("Columna Nombres:", cols_lista, index=idx_nom)
    with col_m2:
        col_ape = st.selectbox("Columna Apellidos:", cols_lista, index=idx_ape)
    with col_m3:
        col_tipo = st.selectbox("Columna Tipo Documento:", cols_lista, index=idx_tipo)
    with col_m4:
        col_doc = st.selectbox("Columna N° Documento:", cols_lista, index=idx_doc)

    df_aprendices["Aprendiz_Info"] = (
        df_aprendices[col_nom].astype(str).str.strip() + " " + 
        df_aprendices[col_ape].astype(str).str.strip() + " | " + 
        df_aprendices[col_tipo].astype(str).str.strip() + ": " + 
        df_aprendices[col_doc].astype(str).str.strip()
    )

    st.write(f"🟢 **{len(df_aprendices)}** Aprendices en estado **'EN FORMACION'** encontrados.")
    st.dataframe(df_aprendices[["Aprendiz_Info"]], use_container_width=True)

    def generar_pdf_aprendiz(row):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        styles = getSampleStyleSheet()
        
        style_title = ParagraphStyle('TitleStyle', parent=styles['Heading2'], alignment=1, fontSize=11, leading=14)
        style_cell = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10)
        style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold")

        elements = []

        # Encabezado SENA
        header_data = [
            [Paragraph("<b>SERVICIO NACIONAL DE APRENDIZAJE SENA<br/>CENTRO NACIONAL COLOMBO ALEMAN<br/>PLAN DE TRABAJO</b>", style_title)]
        ]
        t_header = Table(header_data, colWidths=[540])
        t_header.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke)
        ]))
        elements.append(t_header)
        elements.append(Spacer(1, 5))

        # Información General
        info_data = [
            [Paragraph("<b>Programa de Formación:</b>", style_header), Paragraph(programa, style_cell), Paragraph("<b>Instructor:</b>", style_header), Paragraph(instructor, style_cell)],
            [Paragraph("<b>Proyecto Formativo:</b>", style_header), Paragraph(proyecto, style_cell), Paragraph("<b>Fase del Proyecto:</b>", style_header), Paragraph(fase_proyecto, style_cell)],
            [Paragraph("<b>Nombre del Aprendiz:</b>", style_header), Paragraph(f"{row[col_nom]} {row[col_ape]}", style_cell), Paragraph("<b>Doc. Identidad:</b>", style_header), Paragraph(f"{row[col_tipo]} {row[col_doc]}", style_cell)],
            [Paragraph("<b>Competencia:</b>", style_header), Paragraph(str(competencia_sel), style_cell), Paragraph("<b>Resultado Aprendizaje:</b>", style_header), Paragraph(str(rap_sel), style_cell)]
        ]
        t_info = Table(info_data, colWidths=[110, 160, 110, 160])
        t_info.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
        ]))
        elements.append(t_info)
        elements.append(Spacer(1, 10))

        # Tabla de Actividades (1 a 10)
        actividades_data = [
            [Paragraph("<b>N°</b>", style_header), Paragraph("<b>Actividad a Desarrollar</b>", style_header), Paragraph("<b>Forma Entrega</b>", style_header), Paragraph("<b>¿Entregó?</b>", style_header)]
        ]

        entrega_val = "" if tipo_reporte.startswith("Inicial") else "SI"

        for i in range(1, 11):
            actividades_data.append([
                Paragraph(str(i), style_cell),
                Paragraph(f"Actividad de aprendizaje / Evidencia N° {i}", style_cell),
                Paragraph(forma_entrega, style_cell),
                Paragraph(entrega_val, style_cell)
            ])

        t_act = Table(actividades_data, colWidths=[30, 330, 90, 90])
        t_act.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2EFDA")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
        ]))
        elements.append(t_act)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    st.markdown("---")
    if st.button("🚀 Generar PDF Único de Todos los Aprendices"):
        if df_aprendices.empty:
            st.warning("No hay aprendices en estado 'EN FORMACION' para generar.")
            st.stop()

        pdf_merger = PdfWriter()

        with st.spinner("Generando y uniendo documentos PDF..."):
            for _, row in df_aprendices.iterrows():
                pdf_buf = generar_pdf_aprendiz(row)
                reader = PdfReader(pdf_buf)
                for page in reader.pages:
                    pdf_merger.add_page(page)

        output_pdf = io.BytesIO()
        pdf_merger.write(output_pdf)
        output_pdf.seek(0)

        st.success(f"¡Se generó correctamente el PDF consolidado con los **{len(df_aprendices)}** aprendices!")
        
        st.download_button(
            label="📥 Descargar Documento PDF Único Consolidado",
            data=output_pdf,
            file_name="Planes_Concertados_Consolidado.pdf",
            mime="application/pdf"
        )
