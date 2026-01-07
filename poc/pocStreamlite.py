import streamlit as st
import pandas as pd
from docx import Document
import os

st.title("Briefgenerator met voorgeladen bestanden")

# Standaardbestandsnamen
default_excel = "testadressen.xlsx"
default_template = "BriefSjabloon.docx"

# Controleer of de standaardbestanden bestaan
excel_exists = os.path.exists(default_excel)
template_exists = os.path.exists(default_template)

# Laad standaardbestanden of vraag om upload
if excel_exists and template_exists:
    st.success("Standaardbestanden gevonden! Je kunt ze gebruiken of andere uploaden.")
    df = pd.read_excel(default_excel)
    uploaded_template = default_template
else:
    st.warning("Standaardbestanden niet gevonden. Upload ze hieronder.")
    uploaded_excel = st.file_uploader("Upload Excel-bestand", type=["xlsx"])
    uploaded_template = st.file_uploader("Upload Word-sjabloon", type=["docx"])
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel)
    else:
        st.stop()

# Als standaardbestanden wel bestaan, maar je wilt ze overschrijven
uploaded_excel = st.file_uploader("Upload een ander Excel-bestand (optioneel)", type=["xlsx"], key="excel_uploader")
uploaded_template_file = st.file_uploader("Upload een ander Word-sjabloon (optioneel)", type=["docx"], key="template_uploader")

# Gebruik de geüploade bestanden als ze beschikbaar zijn
if uploaded_excel is not None:
    df = pd.read_excel(uploaded_excel)
if uploaded_template_file is not None:
    uploaded_template = uploaded_template_file.name

# Toon afdelingen en laat gebruiker kiezen
afdelingen = df["Afdeling"].tolist()
geselecteerde_afdeling = st.selectbox("Kies een afdeling", afdelingen)

if geselecteerde_afdeling:
    afdeling_data = df[df["Afdeling"] == geselecteerde_afdeling].iloc[0].to_dict()

    # Vul het sjabloon
    doc = Document(uploaded_template)
    for paragraph in doc.paragraphs:
        for key, value in afdeling_data.items():
            if f"[{key}]" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"[{key}]", str(value))

    # Sla het gevulde document op
    output_path = "gevuld_document.docx"
    doc.save(output_path)

    # Downloadknop
    with open(output_path, "rb") as f:
        st.download_button(
            label="Download gevuld document",
            data=f,
            file_name=output_path,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
