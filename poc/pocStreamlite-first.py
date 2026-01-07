import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# Laad het Excel-bestand
@st.cache_data
def load_data(file):
    return pd.read_excel(file)

# Vul het Word-sjabloon
def fill_template(template_path, output_path, data):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if f"[{key}]" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"[{key}]", str(value))
    doc.save(output_path)

# Streamlit UI
st.title("Briefgenerator")
uploaded_excel = st.file_uploader("Upload Excel-bestand", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Word-sjabloon", type=["docx"])

if uploaded_excel and uploaded_template:
    df = load_data(uploaded_excel)
    afdelingen = df["Afdeling"].tolist()
    geselecteerde_afdeling = st.selectbox("Kies een afdeling", afdelingen)

    if geselecteerde_afdeling:
        afdeling_data = df[df["Afdeling"] == geselecteerde_afdeling].iloc[0].to_dict()

        # Vul het sjabloon
        output_path = "gevuld_document.docx"
        fill_template(uploaded_template.name, output_path, afdeling_data)

        # Downloadknop
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download gevuld document",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
