# streamlit_app.py
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
    "El Patrocinador y/o el médico del estudio garantizará su acceso a este método anticonceptivo "
    "acordado y necesario para su participación en el ensayo. El costo de los métodos anticonceptivos "
    "seleccionados correrá a cargo del Patrocinador."
)

# -----------------------------
# Funciones auxiliares
# -----------------------------
def remove_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def replace_text_in_runs(paragraph, old, new):
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)

def replace_text_in_doc(doc, replacements):
    # Reemplazo en párrafos
    for p in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_in_runs(p, old, new)
        # fallback si quedó texto entero
        fulltext = p.text
        for old, new in replacements.items():
            if old in fulltext:
                for r in p.runs:
                    r.text = ""
                p.add_run(fulltext.replace(old, new))
    # Reemplazo en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in replacements.items():
                        replace_text_in_runs(p, old, new)
                    fulltext = p.text
                    for old, new in replacements.items():
                        if old in fulltext:
                            for r in p.runs:
                                r.text = ""
                            p.add_run(fulltext.replace(old, new))

def find_paragraphs_containing(doc, snippet):
    res = []
    for p in doc.paragraphs:
        if snippet.lower() in p.text.lower():
            res.append(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if snippet.lower() in p.text.lower():
                        res.append(p)
    return res

def process_row_and_generate_doc(template_bytes, row):
    doc = Document(io.BytesIO(template_bytes))

    # -----------------------------
    # Mapeo de placeholders <<...>> con columnas del Excel
    # -----------------------------
    replacements = {
        "<<NUMERO_PROTOCOLO>>": str(row.get("Numero de protocolo", "")).strip(),
        "<<TITULO_ESTUDIO>>": str(row.get("Titulo del Estudio", "")).strip(),
        "<<PATROCINADOR>>": str(row.get("Patrocinador", "")).strip(),
        "<<INVESTIGADOR>>": str(row.get("Investigador", "")).strip(),
        "<<INSTITUCION>>": str(row.get("Institucion", "")).strip(),
        "<<DIRECCION>>": str(row.get("Direccion", "")).strip(),
        "<<CARGO_INVESTIGADOR>>": str(row.get("Cargo del Investigador en la Institucion", "")).strip(),
        "<<Centro_Nro.>>": str(row.get("Nro. de Centro", "")).strip(),
        "<<COMITE>>": str(row.get("COMITE", "")).strip(),
        "<<SUBINVESTIGADOR>>": str(row.get("Subinvestigador", "")).strip(),
        "<<TELEFONO_24HS>>": str(row.get("TELEFONO 24HS", "")).strip(),
        "<<TELEFONO_24HS_SUBINV>>": str(row.get("TELEFONO 24HS subinvestigador", "")).strip(),
    }

    # Reemplazamos placeholders
    replace_text_in_doc(doc, replacements)

    # -----------------------------
    # Lógica de provincia
    # -----------------------------
    prov = str(row.get("provincia", "")).strip().lower()

    if prov == "cordoba":
        paras = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        for p in paras:
            try:
                remove_paragraph(p)
            except Exception:
                pass
        paras_ba = find_paragraphs_containing(doc, "Requerido para centros de la provincia de Buenos Aires")
        for p in paras_ba:
            try:
                remove_paragraph(p)
            except Exception:
                pass

    elif prov.replace(" ", "") in ("buenosaires", "buenosaires"):
        paras = find_paragraphs_containing(doc, texto_anticonceptivo_original)
        if para
