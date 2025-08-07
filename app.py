import streamlit as st
from docx import Document
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

def convert_docx_to_pdf(docx_file):
    doc = Document(docx_file)
    buffer = BytesIO()
    pdf = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()
    content = []

    for para in doc.paragraphs:
        content.append(Paragraph(para.text, styles["Normal"]))

    pdf.build(content)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="Docx to PDF Converter", layout="centered")

st.title("📄 Docx to PDF Converter")
st.write("Upload a `.docx` file to convert it into a PDF.")

uploaded_file = st.file_uploader("Choose a .docx file", type="docx")

if uploaded_file is not None:
    if st.button("Convert to PDF"):
        with st.spinner("Converting..."):
            pdf_buffer = convert_docx_to_pdf(uploaded_file)
            st.success("Conversion successful!")
            st.download_button(
                label="📥 Download PDF",
                data=pdf_buffer,
                file_name="converted.pdf",
                mime="application/pdf"
            )
