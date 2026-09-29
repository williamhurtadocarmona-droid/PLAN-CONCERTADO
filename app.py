import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import date

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄", layout="wide")
st.title("📄 Generador de Planes de Trabajo - SENA")

uploaded_excel = st.file_uploader("Cargar Reporte de Juicios Evaluativos (Excel)", type=["xls", "xlsx"])

if uploaded_excel:
    try:
        prog_sel = st.selectbox("Especialidad/Programa:", ["TOTF - Operación en Torno y Fresadora", "TMMI - Mantenimiento Mecánico Industrial"])
        df_meta = pd.read_excel(uploaded_excel, header=None, nrows=12)
        
        num_ficha = next((str(r[2] if pd.notna(r[2]) else r[1]).strip() for _, r in df_meta.iterrows() if str(r[0]).strip() == "Ficha de Caracterización:"), "")
        
        is_totf = "TOTF" in prog_sel
        den_prog = "OPERACION EN TORNO Y FRESADORA" if is_totf else "MECANICA DE MAQUINARIA INDUSTRIAL"
        proj_def = "OPTIMIZACIÓN EN LA FABRICACIÓN DE COMPONENTES MECÁNICOS EN TORNO Y FRESADORA EN LAS INDUSTRIAS DEL ATLÁNTICO" if is_totf else "IMPLEMENTACIÓN DEL PROGRAMA DE MANTENIMIENTO MECÁNICO INDUSTRIAL EN INDUSTRIAS Y CENTROS DE FORMACION SENA REGIONAL ATLÁNTICO."

        st.info(f"📌 **Programa:** {den_prog} | **Ficha:** {num_ficha}")

        c1, c2 = st.columns(2)
        proj_input = c1.text_input("Proyecto Formativo:", value=proj_def)
        fase_input = c2.selectbox("Fase del Proyecto:", ["Análisis", "Planeación", "Ejecución", "Evaluación"], index=2)

        st.subheader("📅 Fechas del Plan de Trabajo")
        f_col1, f_col2 = st.columns(2)
        fecha_concertada = f_col1.date_input("Fecha Concertada (Inicio):", value=date.today())
        fecha_final = f_col2.date_input("Fecha Final de Entrega:", value=date.today())

        st.subheader("🎯 Competencia y Actividades")
        if is_totf:
            st.selectbox("Competencia:", ["290201211 - MECANIZAR PIEZA INDUSTRIAL DE ACUERDO CON TÉCNICAS MANUALES Y SEMIAUTOMÁTICAS"])
            rap_sel = st.selectbox("Resultado de Aprendizaje (RAP):", [
                "694494 - 1.OPERAR TORNO CONVENCIONAL DE ACUERDO CON PROCEDIMIENTOS TÉCNICOS Y NORMATIVA.",
                "694495 - 2.OPERAR FRESADORA CONVENCIONAL DE ACUERDO CON PROCEDIMIENTOS TÉCNICOS Y NORMATIVA"
            ])
            acts_base = [
                "Preparación de herramientas y equipo: Identificar el tipo de herramienta de corte según material y operación. Acondicionar y afilar herramientas aplicando ángulos adecuados. Seleccionar elementos de sujeción.",
                "Elaboración de la orden operacional: Analizar plano técnico de la pieza, secuenciar operaciones de mecanizado y calcular velocidades y avances.",
                "Puesta a punto del torno convencional: Montar la pieza, verificar sujeción y alineación, instalar herramienta y calibrar alturas.",
                "Ejecución de operaciones de torneado: Cilindrado, refrentado, taladrado, roscado y tronzado según plano técnico."
            ] if "694494" in rap_sel else [
                "Preparación y puesta a punto de máquina y herramientas: Seleccionar cortadores y montar pieza con sistemas de sujeción.",
                "Ejecución del fresado según orden operacional: Interpretar plano y definir secuencia de operaciones.",
                "Control dimensional y verificación: Medir cotas de la pieza fresada con instrumentos de metrología.",
                "Identificación de fallas y mejoras: Detectar defectos y reportar en formato técnico.",
                "Seguridad y protección en fresado: Usar EPP y buenas prácticas en manipulación de virutas."
            ]
        else:
            st.selectbox("Competencia:", ["REPARAR EQUIPOS SEGUN PROCEDIMIENTOS Y MANUALES TECNICOS."])
            rap_sel = st.selectbox("Resultado de Aprendizaje (RAP):", ["OPERAR MÁQUINAS Y HERRAMIENTAS CONVENCIONALES SEGÚN ESPECIFICACIONES TÉCNICAS."])
            acts_base = ["Fabricar elementos mecánicos aplicando procesos de mecanizado con torno convencional."]

        tipo_plan = st.radio("Momento del Plan:", ["Plan Inicial", "Plan Final"], horizontal=True)
        m1, m2 = st.columns(2)
        ent_masiva = m1.selectbox("Forma de Entrega masiva:", ["Personalizar", "Física (Todas)", "Digital (Todas)"])
        est_masiva = m2.selectbox("Estado masivo (Plan Final):", ["Personalizar", "SÍ (Todos)", "NO (Ninguno)"]) if tipo_plan == "Plan Final" else "Personalizar"

        st.subheader("📝 Configuración de Actividades")
        acts_desc, ent_act, est_act = [], [], []
        
        for i, txt in enumerate(acts_base, 1):
            st.markdown(f"--- **Actividad {i}** ---")
            cols = st.columns(3 if tipo_plan == "Plan Final" else 2)
            acts_desc.append(cols[0].text_area(f"Desc {i}", value=txt, height=75, label_visibility="collapsed"))
            ent_act.append(cols[1].radio(f"Ent {i}", ["Física", "Digital"], index=0 if "Física" in ent_masiva else 1, horizontal=True))
            if tipo_plan == "Plan Final":
                est_act.append(cols[2].radio(f"Est {i}", ["SÍ", "NO"], index=0 if "SÍ" in est_masiva else 1, horizontal=True))

        df_raw = pd.read_excel(uploaded_excel, skiprows=12)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        aprendices = df_raw[df_raw['Estado'].str.upper() == 'EN FORMACION'][['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos']].drop_duplicates()

        st.success(f"✅ Aprendices detectados: **{len(aprendices)}**")

        if st.button("📦 Generar y Descargar Planes en ZIP (Word Original)", use_container_width=True):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for _, ap in aprendices.iterrows():
                    doc = Document("Plan de trabajo .docx")
                    
                    reemplazos = {
                        "«Nombre»": str(ap['Nombre']), "«Apellidos»": str(ap['Apellidos']),
                        "«Tipo_de_Doc»": str(ap['Tipo de Documento']), "«N_Documento»": str(ap['Número de Documento']),
                        "TECNICO INSTALACION SISTEMAS ELECTRICOS RESIDENCIALES Y COMERCIALES": den_prog,
                        "«Programa»": den_prog, "«Denominacion»": den_prog,
                        "837101": num_ficha, "«Ficha»": num_ficha, "«Numero_Ficha»": num_ficha,
                        "«Proyecto_Formativo»": proj_input, "«Proyecto»": proj_input,
                        "«Fase_Proyecto»": fase_input, "«Fase»": fase_input
                    }
                    
                    for p in doc.paragraphs:
                        for k, v in reemplazos.items():
                            if k in p.text: p.text = p.text.replace(k, v)
                            
                    if doc.tables:
                        t0 = doc.tables[0]
                        for row in t0.rows:
                            for cell in row.cells:
                                for k, v in reemplazos.items():
                                    if k in cell.text: cell.text = cell.text.replace(k, v)
                        
                        for row in t0.rows:
                            for c_idx, cell in enumerate(row.cells):
                                txt_celda = cell.text.strip().upper()
                                if "PROYECTO FORMATIVO" in txt_celda or txt_celda == "PROYECTO FORMATIVO:":
                                    if c_idx + 1 < len(row.cells):
                                        row.cells[c_idx + 1].text = proj_input
                                if "FASE DEL PROYECTO" in txt_celda or "FASE" in txt_celda:
                                    if c_idx + 1 < len(row.cells):
                                        row.cells[c_idx + 1].text = fase_input

                    if len(doc.tables) > 1:
                        t_act = doc.tables[1]
                        f_conc_str = fecha_concertada.strftime("%d/%m/%Y")
                        f_fin_str = fecha_final.strftime("%d/%m/%Y")

                        for idx, desc in enumerate(acts_desc):
                            r_idx = idx + 3
                            if r_idx < len(t_act.rows):
                                cells = t_act.rows[r_idx].cells
                                cells[0].text = rap_sel
                                cells[1].text = str(idx + 1)
                                cells[2].text = desc
                                cells[3].text = "X" if ent_act[idx] == "Física" else ""
                                cells[4].text = "" if ent_act[idx] == "Física" else "X"
                                
                                cells[5].text = f_conc_str
                                cells[6].text = f_fin_str

                                if tipo_plan == "Plan Final":
                                    cells[7].text = "X" if est_act[idx] == "SÍ" else ""
                                    cells[8].text = "X" if est_act[idx] == "NO" else ""
                                else:
                                    cells[7].text = ""
                                    cells[8].text = ""

                        for r_rm in range(12, len(acts_desc) + 2, -1):
                            if r_rm < len(t_act.rows):
                                tr = t_act.rows[r_rm]._tr
                                tr.getparent().remove(tr)

                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    sigla = prog_sel.split(' - ')[0]
                    zip_file.writestr(f"Plan_Trabajo_{sigla}_{ap['Número de Documento']}_{ap['Nombre']}_{ap['Apellidos']}.docx", doc_io.getvalue())

            st.download_button("⬇️ Descargar Archivo ZIP", data=zip_buffer.getvalue(), file_name=f"Planes_Trabajo_{prog_sel.split(' - ')[0]}_{tipo_plan.replace(' ', '_')}_Ficha_{num_ficha}.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Error al procesar: {e}")
