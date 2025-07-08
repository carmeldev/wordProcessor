import streamlit as st
import os
import tempfile
from bodyParser import process_document

st.set_page_config(page_title="Document Processor", layout="centered")
st.title("📄 Document Formatter & PDF Converter")

# File Uploads
uploaded_docx = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
uploaded_logo = st.file_uploader("Upload Logo (.png/.jpg)", type=["png", "jpg", "jpeg"])

# --- Header/Footer Fields ---
st.subheader("Header & Footer Info")

line1 = st.text_input("Header Line 1 (e.g. ISSN)", value="ISSN (Online): 2455-3662")
line2 = st.text_input("Header Line 2 (Journal Name)", value="EPRA International Journal of Multidisciplinary Research (IJMR) - Peer Reviewed Journal")
line3 = st.text_input("Header Line 3 (Volume/Issue/etc)", value="Volume:11 | Issue:6 | June 2025 || Journal DOI: 10.36713/epra2013 || SJIF Impact Factor 2025: 8.691 || ISI Value: 1.188")

start_page_number = st.number_input("Start Page Number", min_value=1, value=3)
doi_url = st.text_input("DOI URL (optional)", value="https://doi.org/10.36713/epra2013")
footer_journal = st.text_input("Footer Journal Name", value="EPRA IJMR")

journal_code = st.selectbox("Journal Code (for style settings)", ["IJMR", "EPRA", "Custom"], index=0)

# --- Submit and Process ---
if uploaded_docx and uploaded_logo:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.docx")
        logo_path = os.path.join(tmpdir, "logo.png")
        output_docx_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "wb") as f:
            f.write(uploaded_docx.read())

        with open(logo_path, "wb") as f:
            f.write(uploaded_logo.read())

        with st.spinner("Processing document..."):
            process_document(
                input_doc=input_path,
                logo_path=logo_path,
                output_doc=output_docx_path,
                journalCode=journal_code,
                line1=line1,
                line2=line2,
                line3=line3,
                start_page_number=start_page_number,
                doi_url=doi_url,
                footer_journal=footer_journal
            )

            output_pdf_path = output_docx_path.replace(".docx", ".pdf")

        st.success("✅ Document processed successfully!")

        with open(output_docx_path, "rb") as f:
            st.download_button("📄 Download DOCX", f, file_name="formatted_output.docx")

        if os.path.exists(output_pdf_path):
            with open(output_pdf_path, "rb") as f:
                st.download_button("📄 Download PDF", f, file_name="formatted_output.pdf")

else:
    st.info("Please upload both a .docx file and a logo image.")
