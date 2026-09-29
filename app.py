import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import re

st.set_page_config(page_title="Generador de Planes de Trabajo SENA", page_icon="📄", layout="wide")

st.title("📄 Generador de Planes de Trabajo - SENA")
st.write("Sube el archivo de **Reporte de Juicios Evaluativos (.xls/.xlsx)**, selecciona el **Programa**, configura los **Resultados de Aprendizaje**, **Actividades**, **Forma de Entrega** y el **Estado del Plan**.")

# Cargar automáticamente la Planeación Pedagógica si existe en el proyecto
@st.cache_data
def cargar_mapa_planeacion():
    totf_map = {}
    try:
        excel_path = "GPFI-F-134V05Formatoplaneacionpedagogica 2026.xlsx"
        xls = pd.ExcelFile(excel_path)
        for sheet in xls.sheet_names:
            if 'TRIMESTRE' in sheet.upper():
                df_p = pd.read_excel(excel_path, sheet_name=sheet)
                header_row = -1
                for r in range(10, min(20, len(df_p))):
                    row_str = " ".join([str(x).strip().upper() for x in df_p.iloc[r].values if pd.notna(x)])
                    if 'RESULTADOS DE APRENDIZAJE' in row_str:
                        header_row = r
                        break
                if header_row != -1:
                    rap_col, act_col = -1, -1
                    for c in range(df_p.shape[1]):
                        val = str(df_p.iloc[header_row, c]).strip().upper()
                        if 'RESULTADOS DE APRENDIZAJE' in val:
                            rap_col = c
                        if 'ACTIVIDADES DE APRENDIZAJE' in val:
                            act_col = c
                    
                    curr_rap = ""
                    for r in range(header_row + 1, len(df_p)):
                        val_rap = str(df_p.iloc[r, rap_col]).strip() if pd.notna(df_p.iloc[r, rap_col]) else ""
                        val_act = str(df_p.iloc[r, act_col]).strip() if pd.notna(df_p.iloc[r, act_col]) else ""
                        if val_rap and val_rap != 'nan':
                            curr_rap = " ".join(val_rap.split())
                        if curr_rap and val_act and val_act != 'nan':
                            clean_act = " ".join(val_act.split())
                            totf_map[curr_rap] = clean_act
    except Exception as e:
        pass
    return totf_map

totf_actividades_map = cargar_mapa_planeacion()

# Subir archivo Excel desde la interfaz web
uploaded_excel = st.file_uploader("Cargar Reporte de Juicios Evaluativos (Excel)", type=["xls", "xlsx"])

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
        
        denominacion_programa = ""
        numero_ficha = ""
        
        for idx, row in df_meta.iterrows():
            label = str(row[0]).strip()
            if label == "Ficha de Caracterización:":
                numero_ficha = str(row[2]).strip() if pd.notna(row[2]) else str(row[1]).strip()
            elif label == "Denominación:":
                denominacion_programa = str(row[2]).strip() if pd.notna(row[2]) else str(row[1]).strip()

        # Si se selecciona TOTF, asigna valores predeterminados de TOTF
        if "TOTF" in programa_seleccionado:
            denominacion_programa = "OPERACION EN TORNO Y FRESADORA"
            proyecto_default = "OPTIMIZACIÓN EN LA FABRICACIÓN DE COMPONENTES MECÁNICOS EN TORNO Y FRESADORA EN LAS INDUSTRIAS DEL ATLÁNTICO"
        else:
            proyecto_default = ""

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
                help="Se autocompleta cuando seleccionas TOTF."
            )
            
        with col_proj2:
            fase_proyecto_input = st.selectbox(
                "Fase del Proyecto:",
                options=["Análisis", "Planeación", "Ejecución", "Evaluación"],
                index=2, # Ejecución por defecto
                help="Selecciona la Fase del Proyecto correspondiente."
            )

        # 6. Obtener lista de Resultados de Aprendizaje (RAPs) disponibles
        raps_disponibles = sorted(df_filtrado['Resultado de Aprendizaje'].dropna().unique().tolist())
        
        st.subheader("🎯 2. Selección de Resultados de Aprendizaje (RAP)")
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
            # 7. Opciones del Tipo de Plan (Inicial vs Final)
            st.subheader("📋 3. Estado del Plan de Trabajo")
            tipo_plan = st.radio(
                "Selecciona el momento de generación del Plan de Trabajo:",
                options=["Plan Inicial", "Plan Final"],
                index=0,
                horizontal=True,
                help="En 'Plan Inicial' los estados se completan automáticamente según el Excel. En 'Plan Final' puedes definir la entrega manual de cada actividad."
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

            # 9. Configurar Actividades a desarrollar, Forma de Entrega y Estado de Entrega
            st.subheader("📝 5. Configurar Actividades e Individualizar")
            
            actividades_por_rap = {}
            entrega_por_rap = {}
            estado_entrega_final = {}

            idx_entrega_default = 0 if "Física" in entrega_masiva else 1
            idx_estado_default = 0 if "SÍ" in estado_masivo else 1

            # Función para buscar la actividad predeterminada según el RAP
            def obtener_actividad_predeterminada(rap_str):
                if "TOTF" in programa_seleccionado and totf_actividades_map:
                    # Búsqueda exacta primero
                    if rap_str in totf_actividades_map:
                        return totf_actividades_map[rap_str]
                    # Búsqueda por código de 6 dígitos
                    m = re.search(r'\d{6}', rap_str)
                    if m:
                        code = m.group(0)
                        for k_map, v_map in totf_actividades_map.items():
                            if code in k_map:
                                return v_map
                return f"Desarrollar guía de aprendizaje y evidencias prácticas de: {rap_str.split('-')[-1].strip()}"

            for i, rap in enumerate(raps_seleccionados, 1):
                st.markdown("---")
                st.markdown(f"**Actividad {i}:** `{rap}`")
                
                if tipo_plan == "Plan Final":
                    col1, col2, col3 = st.columns([3, 1, 1])
                else:
                    col1, col2 = st.columns([3, 1])
                
                default_act_val = obtener_actividad_predeterminada(rap)
                
                with col1:
                    actividades_por_rap[rap] = st.text_area(
                        f"Descripción de la Actividad {i}",
                        value=default_act_val,
                        key=f"act_rap_{i}",
                        height=100
                    )
                
                with col2:
                    sub_idx_entrega = idx_entrega_default if "Todas" in entrega_masiva else 1
                    entrega_por_rap[rap] = st.radio(
                        f"Forma de Entrega {i}",
                        options=["Física", "Digital"],
                        index=sub_idx_entrega,
                        key=f"entrega_rap_{i}"
                    )
                
                if tipo_plan == "Plan Final":
                    with col3:
                        sub_idx_estado = idx_estado_default if ("Todos" in estado_masivo or "Ninguno" in estado_masivo) else 0
                        estado_entrega_final[rap] = st.radio(
                            f"¿Entregó Actividad {i}?",
                            options=["SÍ", "NO"],
                            index=sub_idx_estado,
                            key=f"estado_entrega_{i}"
                        )

            # 10. Obtener la lista de aprendices únicos en formación
            aprendices = df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombre', 'Apellidos', 'Estado']].drop_duplicates()
            st.success(f"✅ Aprendices a procesar: **{len(aprendices)}** | Programa: **{programa_seleccionado.split(' - ')[0]}** | RAPs a evaluar: **{num_raps}**")
            
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
                            
                            # Proyecto Formativo y Fase
                            "«Proyecto_Formativo»": proyecto_formativo_input,
                            "«Proyecto»": proyecto_formativo_input,
                            "«Fase_Proyecto»": fase_proyecto_input,
                            "«Fase»": fase_proyecto_input,
                        }
                        
                        # Reemplazar encabezados en párrafos
                        for p in doc.paragraphs:
                            for k, v in reemplazos.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, v)
                                    
                        # Reemplazar encabezados en la primera tabla
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
                        
                        sigla_prog = programa_seleccionado.split(' - ')[0]
                        filename = f"Plan_Trabajo_{sigla_prog}_{num_doc}_{aprendiz['Nombre']}_{aprendiz['Apellidos']}.docx"
                        zip_file.writestr(filename, doc_io.getvalue())
                
                # Botón para descargar el ZIP resultante
                sigla_prog = programa_seleccionado.split(' - ')[0]
                st.download_button(
                    label="📦 Descargar Documentos de Aprendices (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Planes_Trabajo_{sigla_prog}_{tipo_plan.replace(' ', '_')}_Ficha_{numero_ficha}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
