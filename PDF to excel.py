import streamlit as st
import pdfplumber
import re

from openpyxl import load_workbook
from io import BytesIO
from duckduckgo_search import DDGS
def get_company_info_from_web(company_name):
    if not company_name:
        return ""

    query = f"{company_name} company business overview"
    results_text = ""

    with DDGS() as ddgs:
        for r in ddgs.text(query, max_results=5):
            results_text += r.get("body", "") + " "

    return results_text.strip()

def create_short_company_summary(text, max_sentences=4):
    if not text:
        return ""

    sentences = text.split(". ")
    return ". ".join(sentences[:max_sentences]).strip()


st.set_page_config(page_title="PDF → Excel Mapper", layout="centered")

st.title("📄➡️📊 PDF to Pre-Made Excel Template")
st.write("Upload a PDF and an Excel template. Only required fields will be updated.")

# ---------------------------
# PDF TEXT EXTRACTION
# ---------------------------
def extract_pdf_text(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

# ---------------------------
# FIELD EXTRACTION (EDIT THIS)
# ---------------------------
def extract_fields_from_text(text):
    """
    Adjust regex patterns to match YOUR PDFs
    """
    fields = {}

    patterns = {
    "Company_name": r"Company Name[:\s]+(.+)",
    "Org_number": r"Org(?:anisation)? Number[:\s]+([\d\-]+)",
    "Address": r"Address[:\s]+(.+)",
    "Post_nr": r"Post(?:al)? Code[:\s]+(\d+)",
    "city": r"City[:\s]+(.+)",
    "NACE-kode": r"NACE(?: Code)?[:\s]+([\d\.]+)",
    "Omsetning_2024": r"Turnover 2024[:\s]+([\d\s,\.]+)",
    "Hjemmeside": r"(?:Website|Homepage)[:\s]+(\S+)",
    "Number_of_Employees": r"(?:Employees|Number of Employees)[:\s]+(\d+)",
    "email": r"Email[:\s]+([\w\.-]+@[\w\.-]+)"
}


    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        fields[key] = match.group(1).strip() if match else ""

    return fields

# ---------------------------
# EXCEL UPDATE LOGIC
# ---------------------------
def update_excel(template_file, extracted_data, company_summary):

    wb = load_workbook(template_file)
    ws = wb.active  # CHANGE SHEET NAME IF NEEDED

    # 🔴 MAP DATA → CELLS (EDIT THIS)
    cell_mapping = {
        "Company_name": "B14",
        "Org_number": "B15",
        "Address": "B16",
        "Post-nr": "B17",
        "NACE-kode": "B18",
        "Omsetning_2024": "B19",
        "Hjemmeside": "B21",
        "Number_of_Employees": "B22",
    }

    for field, cell in cell_mapping.items():
        if extracted_data.get(field):
            ws[cell] = extracted_data[field]

    # Kort info om företaget
if company_summary:
    ws["B10"] = f"Kort info om företaget:\n{company_summary}"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------------------------
# STREAMLIT UI
# ---------------------------
pdf_file = st.file_uploader("Upload PDF", type="pdf")
excel_file = st.file_uploader("Upload Excel Template", type=["xlsx"])

if pdf_file and excel_file:
    if st.button("Extract & Update Excel"):
        with st.spinner("Processing..."):

            # STEP 1 — PDF extraction
            pdf_text = extract_pdf_text(pdf_file)
            extracted_fields = extract_fields_from_text(pdf_text)

            # STEP 2 — Web search
            company_name = extracted_fields.get("company_name", "")
            web_text = get_company_info_from_web(company_name)
            company_summary = create_short_company_summary(web_text)

            # STEP 3 — Write to Excel
            updated_excel = update_excel(
                excel_file,
                extracted_fields,
                company_summary
            )

        st.success("Excel updated successfully!")

        st.subheader("Extracted data")
        st.json(extracted_fields)

        st.download_button(
            label="Download updated Excel file",
            data=updated_excel,
            file_name="updated_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

