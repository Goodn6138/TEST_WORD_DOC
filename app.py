import streamlit as st
from io import BytesIO
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from doc2docx import convert
import os
import tempfile
import uuid

def convert_doc_to_docx(doc_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as tmp_doc:
        tmp_doc.write(doc_file.read())
        doc_path = tmp_doc.name

    output_docx = doc_path + ".docx"
    convert(doc_path, output_docx)
    return output_docx

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def text_to_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()
    content = [Paragraph(line, styles["Normal"]) for line in text.splitlines()]
    doc.build(content)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="Word to PDF Converter", layout="centered")

st.title("📄 Word (.doc/.docx) to PDF Converter")
uploaded_file = st.file_uploader("Upload a Word file", type=["doc", "docx"])

if uploaded_file and st.button("Convert to PDF"):
    with st.spinner("Processing..."):
        try:
            suffix = os.path.splitext(uploaded_file.name)[1].lower()

            if suffix == ".doc":
                docx_path = convert_doc_to_docx(uploaded_file)
                text = extract_text_from_docx(docx_path)
                os.remove(docx_path)  # Clean up
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(uploaded_file.read())
                    tmp.flush()
                    text = extract_text_from_docx(tmp.name)

            pdf_buffer = text_to_pdf(text)
            st.success("✅ Conversion successful!")
            st.download_button(
                label="📥 Download PDF",
                data=pdf_buffer,
                file_name="converted.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ Error: {e}")
