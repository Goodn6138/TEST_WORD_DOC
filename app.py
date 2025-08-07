import streamlit as st
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import textract
import tempfile

def extract_text(file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=file.name) as tmp:
        tmp.write(file.read())
        tmp_path = tmp.name

    try:
        text = textract.process(tmp_path).decode("utf-8")
        return text
    except Exception as e:
        return None

def convert_text_to_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()
    flowables = []

    for line in text.splitlines():
        flowables.append(Paragraph(line, styles["Normal"]))

    doc.build(flowables)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="Doc/Docx to PDF Converter", layout="centered")

st.title("📄 Word to PDF Converter")
st.write("Upload a `.doc` or `.docx` file and get a PDF.")

uploaded_file = st.file_uploader("Choose a file", type=["doc", "docx"])

if uploaded_file:
    if st.button("Convert to PDF"):
        with st.spinner("Extracting and converting..."):
            extracted_text = extract_text(uploaded_file)
            if extracted_text:
                pdf_buffer = convert_text_to_pdf(extracted_text)
                st.success("✅ Conversion successful!")
                st.download_button(
                    label="📥 Download PDF",
                    data=pdf_buffer,
                    file_name="converted.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("❌ Failed to extract text from file.")
