import os
import streamlit as st
import pandas as pd
from pptx import Presentation
import subprocess

st.set_page_config(page_title="📄 Report Card Generator", layout="wide")

st.title("📄 Report Card Generator")
st.write("Upload an **Excel file** and a **PowerPoint template** to generate personalized report cards.")

excel_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])
pptx_file = st.file_uploader("📂 Upload PowerPoint Template", type=["pptx"])

output_dir = "Generated_Reports"
os.makedirs(output_dir, exist_ok=True)

def replace_text(slide, replacements):
    """Replace placeholders in PPTX."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    for key, value in replacements.items():
                        if key in run.text:
                            run.text = run.text.replace(key, value)

def convert_ppt_to_pdf(ppt_path, pdf_path):
    """Convert PPTX to PDF using LibreOffice."""
    try:
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(pdf_path), ppt_path
        ], check=True)
        print(f"✅ Converted {ppt_path} to {pdf_path}")
    except Exception as e:
        print(f"❌ Error converting {ppt_path} to PDF: {e}")

if excel_file and pptx_file:
    df = pd.read_excel(excel_file)

    for _, row in df.iterrows():
        prs = Presentation(pptx_file)
        slide = prs.slides[0]

        replacements = {
            "{{Student Name}}": str(row["Student Name"]),
            "{{School Name}}": str(row["School Name"]),
            "{{Grade}}": str(row["Grade"]),
            "{{Roll No.}}": str(row["Roll No."]),
            "{{Academic year}}": str(row["Academic Year"]),
            "{{Date of Issue}}": str(row["Date of Issue"]),
            "{{Assessment Grade}}": str(row["Assessment Grade"])
        }

        replace_text(slide, replacements)

        student_name = row['Student Name'].replace(" ", "_")
        pptx_path = os.path.join(output_dir, f"{student_name}_Report_Card.pptx")
        pdf_path = pptx_path.replace(".pptx", ".pdf")

        prs.save(pptx_path)
        convert_ppt_to_pdf(pptx_path, pdf_path)

        st.success(f"✅ {row['Student Name']}'s Report Card Generated!")
        with open(pdf_path, "rb") as pdf_file:
            st.download_button("📥 Download Report Card (PDF)", pdf_file, file_name=f"{row['Student Name']}_Report_Card.pdf", mime="application/pdf")
