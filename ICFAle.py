# ICFAle.py
import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import re

st.set_page_config(page_title="Generador DOCX Consentimientos", layout="wide")

st.title("Generador automático de Consentimientos (Excel → Word)")

st.markdown("""
Subí tu **modelo.docx** (plantilla con placeholders `<<...>>`) y el **datos.xlsx** con la información de cada investigador.  
El nombre del archivo final incluirá el número de protocolo y el nombre del investigador.
""")

uploaded_docx = st.file_uploader("Subí el modelo (.docx)", type=["docx"])
uploaded_xlsx = st.file_uploader("Subí el Excel (.xlsx)", type=["xlsx"])

# Textos para la lógica de provincia
texto_anticonceptivo_original = (
    "El médico del estudio discutirá con usted qué método anticonceptivo se considera adecuado. "
    "El patrocinador y/o el investigador del estudio garantizarán su acceso al método anticonceptivo "
    "acordado y necesario para su participación en este estudio"
)

texto_ba_reemplazo = (
    "El médico del estudio discutirá con usted qué métodos anticonceptivos se consideran adecuados. "
    "El Patrocinador y/o
