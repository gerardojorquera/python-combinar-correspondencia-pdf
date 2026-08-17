import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, KeepTogether, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import io
import zipfile
from PIL import Image
from PyPDF2 import PdfMerger
import unicodedata

# Configuración inicial de la interfaz web
st.set_page_config(page_title="Generador Masivo de Certificados", page_icon="🎓", layout="wide")

st.title("🎓 Generador Masivo de Certificados de Informática")
st.markdown("Carga tu nómina en Excel, añade los logos corporativos y genera tus diplomas listos para imprimir de forma automatizada.")

# --- SECCIÓN DE AYUDA / NUEVO ESQUEMA SIMPLIFICADO ---
with st.expander("📋 Ver nombres simplificados de las columnas del Excel", expanded=False):
    st.markdown("""
    Tu archivo Excel **debe contener exactamente los siguientes nombres de columna** (puedes escribirlos en mayúsculas o minúsculas):
    * `titulo-certificado` — *Ej: Diploma de Participación*
    * `nombre-participante` — *Ej: Francisco Reyes Retamal*
    * `texto-central` — *Ej: Por su participación en el programa ejecutivo...*
    * `fecha-emision` — *Ej: Santiago de Chile - 13 de agosto de 2026*
    * `nombre-relator` — *Ej: Juan Pérez*
    * `empresa-relator` — *Ej: Relator - NETCapacitaciones*
    * `nombre-coordinador` — *Ej: Carlos Gómez*
    * `empresa-coordinador` — *Ej: Jefe de RRHH - MediSoft*
    """)

# --- PANEL DE CONFIGURACIÓN LATERAL ---
st.sidebar.header("🛠️ Parámetros de Personalización")

# Carga de Logos
st.sidebar.subheader("🖼️ Logos de Identidad")
logo_dicta = st.sidebar.file_uploader("Logo Empresa que Dictó (Top Izquierdo)", type=["png", "jpg", "jpeg"])
logo_recibe = st.sidebar.file_uploader("Logo Empresa que Recibió (Top Derecho)", type=["png", "jpg", "jpeg"])

# Configuración del PDF
st.sidebar.subheader("⚙️ Configuración del Output")
modo_generacion = st.sidebar.radio(
    "Formato de Salida:",
    ["Certificados individuales (Archivo ZIP)", "Todos juntos en un solo PDF consolidado"]
)

# --- PANEL CENTRAL: CARGA DE DATOS ---
st.subheader("📁 Carga de Datos")
excel_file = st.file_uploader("Sube tu archivo de datos (.xlsx)", type=["xlsx"])

# --- FUNCIÓN NATIVA PARA GENERAR DIPLOMA ---
def construir_pdf_certificado(row, logo_left_bytes=None, logo_right_bytes=None):
    """Genera la estructura visual de un diploma en orientación horizontal y mantiene el flujo abierto."""
    output_stream = io.BytesIO()
        
    doc = SimpleDocTemplate(
        output_stream,
        pagesize=landscape(letter),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CertTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=32, leading=38, alignment=TA_CENTER, textColor=colors.HexColor('#1A365D'))
    name_style = ParagraphStyle('CertName', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=24, leading=28, alignment=TA_CENTER, textColor=colors.HexColor('#2C5282'))
    body_style = ParagraphStyle('CertBody', parent=styles['Normal'], fontName='Helvetica', fontSize=14, leading=22, alignment=TA_CENTER, textColor=colors.HexColor('#4A5568'))
    date_style = ParagraphStyle('CertDate', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=12, leading=16, alignment=TA_CENTER, textColor=colors.HexColor('#718096'))
    
    sign_name_style = ParagraphStyle('SignName', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=14, alignment=TA_CENTER, textColor=colors.HexColor('#2D3748'))
    sign_sub_style = ParagraphStyle('SignSub', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=12, alignment=TA_CENTER, textColor=colors.HexColor('#718096'))

    story = []
    
    # Encabezado de logos inicializados vacíos
    cell_left = Paragraph("", styles['Normal'])
    cell_right = Paragraph("", styles['Normal'])
    
    if logo_left_bytes:
        img_l = Image.open(io.BytesIO(logo_left_bytes))
        # 🚀 CORRECCIÓN: Se duplica el tamaño límite a 240x100
        img_l.thumbnail((240, 100))
        img_l_stream = io.BytesIO()
        img_l.save(img_l_stream, format="PNG")
        img_l_stream.seek(0)
        cell_left = RLImage(img_l_stream, width=img_l.width, height=img_l.height)
        
    if logo_right_bytes:
        img_r = Image.open(io.BytesIO(logo_right_bytes))
        # 🚀 CORRECCIÓN: Se duplica el tamaño límite a 240x100
        img_r.thumbnail((240, 100))
        img_r_stream = io.BytesIO()
        img_r.save(img_r_stream, format="PNG")
        img_r_stream.seek(0)
        cell_right = RLImage(img_r_stream, width=img_r.width, height=img_r.height)

    # La tabla distribuye el espacio horizontal de forma equilibrada (5 pulgadas por celda)
    header_table = Table([[cell_left, cell_right]], colWidths=[5 * inch, 5 * inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(header_table)
    # 🚀 CORRECCIÓN: Se reduce el espaciado para compensar el crecimiento de los logos
    story.append(Spacer(1, 15))

    story.append(Paragraph(str(row['titulo-certificado']).upper(), title_style))
    story.append(Spacer(1, 25))
    story.append(Paragraph("Otorgado con orgullo a:", ParagraphStyle('Sub', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12, textColor=colors.gray)))
    story.append(Spacer(1, 15))
    story.append(Paragraph(str(row['nombre-participante']), name_style))
    story.append(Spacer(1, 20))
    story.append(Paragraph(str(row['texto-central']), body_style))
    story.append(Spacer(1, 25))
    story.append(Paragraph(str(row['fecha-emision']), date_style))
    story.append(Spacer(1, 50))
    
    firmas_data = [
        [Paragraph("____________________________", sign_name_style), Paragraph("____________________________", sign_name_style)],
        [Paragraph(str(row['nombre-relator']), sign_name_style), Paragraph(str(row['nombre-coordinador']), sign_name_style)],
        [Paragraph(str(row['empresa-relator']), sign_sub_style), Paragraph(str(row['empresa-coordinador']), sign_sub_style)]
    ]
    
    firmas_table = Table(firmas_data, colWidths=[5 * inch, 5 * inch])
    firmas_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    
    story.append(KeepTogether(firmas_table))
    
    def draw_background(canvas, doc):
        canvas.saveState()
        width, height = doc.pagesize
        canvas.setStrokeColor(colors.HexColor('#1A365D'))
        canvas.setLineWidth(4)
        canvas.rect(15, 15, width - 30, height - 30)
        canvas.setStrokeColor(colors.HexColor('#D69E2E'))
        canvas.setLineWidth(1)
        canvas.rect(20, 20, width - 40, height - 40)
        canvas.restoreState()

    # 🚀 CORRECCIÓN CLAVE: El documento se compila e inyecta los datos de forma segura en memoria
    doc.build(story, onFirstPage=draw_background)
    
    # Reposicionar el puntero al inicio para permitir lecturas masivas posteriores
    output_stream.seek(0)
    return output_stream

# BLOQUE Nº2
def normalizar_columna(txt):
    """Limpia espacios, tildes y guiones para un emparejamiento 100% elástico."""
    if not isinstance(txt, str):
        return ""
    txt_norm = unicodedata.normalize('NFD', txt)
    txt_limpio = "".join(c for c in txt_norm if unicodedata.category(c) != 'Mn')
    return txt_limpio.strip().lower().replace(" ", "").replace("_", "-")

# --- CONTROLADOR PRINCIPAL ---
if excel_file:
    try:
        df = pd.read_excel(excel_file)
        
        # Mapeo estricto del sistema usando los nuevos tokens simplificados
        columnas_sistema = [
            "titulo-certificado", "nombre-participante", "texto-central", 
            "fecha-emision", "nombre-relator", "empresa-relator", 
            "nombre-coordinador", "empresa-coordinador"
        ]
        
        # Crear diccionario de búsqueda normalizado para las columnas que el usuario subió
        columnas_reales_normalizadas = {normalizar_columna(col): col for col in df.columns}
        
        columnas_faltantes = []
        columnas_encontradas_mapeo = {}
        
        for col_token in columnas_sistema:
            token_norm = normalizar_columna(col_token)
            if token_norm in columnas_reales_normalizadas:
                columnas_encontradas_mapeo[col_token] = columnas_reales_normalizadas[token_norm]
            else:
                columnas_faltantes.append(col_token)
        
        if columnas_faltantes:
            st.error(f"⚠️ Al archivo Excel le faltan las siguientes columnas: {', '.join(columnas_faltantes)}")
            st.markdown("💡 *Renombra los campos de tu Excel usando nombres cortos como 'titulo-certificado' o 'nombre-participante'.*")
        else:
            # Reestructurar el dataframe con las claves sanitizadas del sistema
            df_renombrado = df.rename(columns={v: k for k, v in columnas_encontradas_mapeo.items()})
            
            st.success(f"📊 ¡Columnas simplificadas validadas con éxito! Se procesarán **{len(df_renombrado)}** diplomas.")
            st.dataframe(df.head(3), use_container_width=True)
            
            iniciar_generacion = st.button("▶️ Iniciar Generación Masiva", type="primary", use_container_width=True)
            
            if iniciar_generacion:
                bytes_left = logo_dicta.getvalue() if logo_dicta else None
                bytes_right = logo_recibe.getvalue() if logo_recibe else None
                
                total_registros = len(df_renombrado)
                barra_progreso = st.progress(0)
                contenedor_estado = st.empty()
                
                if modo_generacion == "Certificados individuales (Archivo ZIP)":
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for idx, row in df_renombrado.iterrows():
                            num_actual = idx + 1
                            contenedor_estado.info(f"⏳ Armando certificado **{num_actual}** de **{total_registros}**: {row['nombre-participante']}")
                            barra_progreso.progress(num_actual / total_registros)
                            
                            pdf_data = construir_pdf_certificado(row, bytes_left, bytes_right)
                            nombre_limpio = "".join(c for c in str(row['nombre-participante']) if c.isalnum() or c in (' ', '_')).rstrip()
                            zip_file.writestr(f"Certificado_{nombre_limpio}.pdf", pdf_data.read())
                    
                    zip_buffer.seek(0)
                    contenedor_estado.success(f"✅ ¡Proceso completado! Se crearon **{total_registros}** diplomas individuales de forma exitosa.")
                    
                    st.download_button(
                        label="📥 Descargar todos los Certificados (.ZIP)",
                        data=zip_buffer,
                        file_name="certificados_informatica.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                else:
                    merger = PdfMerger()
                    
                    for idx, row in df_renombrado.iterrows():
                        num_actual = idx + 1
                        contenedor_estado.info(f"⏳ Indexando al consolidado de impresión: **{num_actual}** de **{total_registros}**")
                        barra_progreso.progress(num_actual / total_registros)
                        
                        single_pdf = construir_pdf_certificado(row, bytes_left, bytes_right)
                        merger.append(single_pdf)
                        
                    output_consolidado = io.BytesIO()
                    merger.write(output_consolidado)
                    output_consolidado.seek(0)
                    
                    contenedor_estado.success(f"✅ ¡Proceso completado! Documento unificado de **{total_registros}** páginas generado con éxito.")
                    
                    st.download_button(
                        label="📥 Descargar PDF Consolidado Único",
                        data=output_consolidado,
                        file_name="nomina_certificados_unificados.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
    except Exception as e:
        st.error(f"❌ Ocurrió un error al procesar la planilla: {str(e)}")
else:
    st.info("💡 Por favor, sube un archivo Excel válido por el panel central para activar los motores de renderizado.")
