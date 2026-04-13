import io
from datetime import datetime
import streamlit as st
import pandas as pd

# Librerías de lectura
try:
    from docx import Document as DocxDocument
except:
    DocxDocument = None

try:
    import pdfplumber
except:
    pdfplumber = None

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except:
    Document = None

APP_TITLE = "UCCuyo · Valorador de Proyectos de Investigación"
APP_VERSION = "v2.0 – valoración flexible"

# ---------------- CRITERIOS ----------------
CRITERIA = {
    "Pertinencia y relevancia": 10,
    "Claridad del problema y objetivos": 10,
    "Originalidad / aporte": 8,
    "Solidez metodológica": 14,
    "Calidad de datos / muestra": 10,
    "Factibilidad y cronograma": 8,
    "Consideraciones éticas": 6,
    "Impacto esperado": 8,
    "Plan de difusión / transferencia": 6,
    "Presupuesto y sostenibilidad": 6,
    "Alineación institucional y normativa": 6,
    "Bibliografía actualizada": 8
}

# ---------------- RESULTADO ----------------
def categorize(p):
    if p >= 70:
        return "Aprobado"
    elif p >= 50:
        return "Aprobado con observaciones"
    elif p >= 30:
        return "Requiere reformulación"
    else:
        return "No aprobado"

# ---------------- PARSE ----------------
def parse_docx(file):
    if DocxDocument is None:
        return ""
    doc = DocxDocument(io.BytesIO(file))
    return "\n".join(p.text for p in doc.paragraphs)

def parse_pdf(file):
    if pdfplumber is None:
        return ""
    text = ""
    with pdfplumber.open(io.BytesIO(file)) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

# ---------------- UI ----------------
def score_ui():
    st.subheader("Evaluación del proyecto")

    scores = {}
    total = 0
    max_total = sum(CRITERIA.values())

    cols = st.columns(2)

    i = 0
    for crit, peso in CRITERIA.items():
        with cols[i % 2]:
            st.markdown(f"**{crit}** (máx {peso})")

            val = st.slider(
                "Puntaje",
                0,
                peso,
                int(peso * 0.7),
                key=crit
            )

            obs = st.text_area(
                "Observaciones",
                key=f"obs_{crit}"
            )

            scores[crit] = (val, obs)
            total += val
            st.divider()

        i += 1

    porcentaje = (total / max_total) * 100

    return scores, porcentaje

# ---------------- EXCEL ----------------
def make_excel(scores, porcentaje, resultado):
    data = []

    for crit, (val, obs) in scores.items():
        data.append({
            "Criterio": crit,
            "Puntaje": val,
            "Observaciones": obs
        })

    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Evaluación")

        resumen = pd.DataFrame([{
            "Resultado": resultado,
            "Porcentaje": round(porcentaje, 2),
            "Fecha": datetime.now()
        }])

        resumen.to_excel(writer, index=False, sheet_name="Resumen")

    return output.getvalue()

# ---------------- WORD ----------------
def make_word(scores, porcentaje, resultado):
    if Document is None:
        return b""

    doc = Document()

    doc.add_heading("Valoración de Proyecto de Investigación", 1)

    doc.add_paragraph(f"Resultado: {resultado}")
    doc.add_paragraph(f"Cumplimiento: {round(porcentaje,2)}%")

    table = doc.add_table(rows=1, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje"
    hdr[2].text = "Observaciones"

    for crit, (val, obs) in scores.items():
        row = table.add_row().cells
        row[0].text = crit
        row[1].text = str(val)
        row[2].text = obs

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# ---------------- APP ----------------
st.set_page_config(page_title=APP_TITLE, layout="wide")

st.title(APP_TITLE)
st.caption("Versión flexible: el evaluador decide el puntaje")

file = st.file_uploader("Subir proyecto (PDF o DOCX)", type=["pdf", "docx"])

if file:
    content = file.read()

    if file.name.endswith(".pdf"):
        text = parse_pdf(content)
    else:
        text = parse_docx(content)

    scores, porcentaje = score_ui()
    resultado = categorize(porcentaje)

    st.markdown(f"## Resultado: **{resultado}**")
    st.markdown(f"### {round(porcentaje,2)} %")

    col1, col2 = st.columns(2)

    with col1:
        excel = make_excel(scores, porcentaje, resultado)
        st.download_button("Descargar Excel", excel, "resultado.xlsx")

    with col2:
        word = make_word(scores, porcentaje, resultado)
        st.download_button("Descargar Word", word, "resultado.docx")
