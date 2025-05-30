import streamlit as st
import os
import tempfile
import shutil
from main import main as run_credit_risk_engine

st.set_page_config(page_title="Credit Risk Appraisal", layout="centered")

st.title("📊 Credit Risk Appraisal System")
st.write("Upload all relevant CA1 documents for analysis. Accepted formats: PDFs and Excel files.")

# File uploader
uploaded_files = st.file_uploader(
    "Upload CA1 documents (ZIP or multiple files)", 
    type=["pdf", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("Analyzing documents..."):
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Save files to a temp CA1 folder
            case_dir = os.path.join(tmp_dir, "CA1_20250422_0001")
            os.makedirs(case_dir, exist_ok=True)

            for file in uploaded_files:
                file_path = os.path.join(case_dir, file.name)
                with open(file_path, "wb") as f:
                    f.write(file.read())

            # Swap your BASE_DIR to point here temporarily
            os.makedirs("data/ca1_cases", exist_ok=True)
            shutil.rmtree("data/ca1_cases/CA1_20250422_0001", ignore_errors=True)
            shutil.copytree(case_dir, "data/ca1_cases/CA1_20250422_0001")


            # Run your backend logic
            run_credit_risk_engine()

            st.success("✅ Processing complete. Check the results folder.")
