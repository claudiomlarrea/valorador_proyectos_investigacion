import io
from datetime import datetime

import pandas as pd
import streamlit as st

# Lectura de archivos
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

try:
    from docx import Document
except Exception:
    Document = None


APP_TITLE = "UCCuyo · Valorador de Proyectos de Investigación"
APP_VERSION = "v3.0 estable – evaluación manual"

CRITERIOS = {
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
    "Bibliografía actualizada": 8,
}


def categoria(porcentaje: float) -> str:
    if porcentaje >= 70:
        return "Aprobado"
    elif porcentaje >= 50:
        return "Aprobado con observaciones"
    elif porcentaje >= 30:
        return "Requiere reformulación"
    return "No aprobado"


def parse_pdf(file_bytes: bytes) -> str:
    if pdfplumber is None:
        return ""
    partes = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            partes.append(page.extract_text() or "")
    return "\n".join(partes)


def parse_docx(file_bytes: bytes) -> str:
    if DocxDocument is None:
        return ""
    doc = DocxDocument(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs)


def make_excel(scores: dict, porcentaje: float, resultado: str, nombre_archivo: str) -> bytes:
    filas = []
    for criterio, puntaje in scores.items():
        filas.append({
            "Criterio": criterio,
            "Puntaje": puntaje
        })

    df = pd.DataFrame(filas)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")

        resumen = pd.DataFrame([{
            "Archivo": nombre_archivo,
            "Resultado": resultado,
            "Porcentaje": round(porcentaje, 2),
            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M")
        }])
        resumen.to_excel(writer, index=False, sheet_name="Resumen")

    return output.getvalue()


def make_word(scores: dict, porcentaje: float, resultado: str, nombre_archivo: str) -> bytes:
    if Document is None:
        return b""

    doc = Document()
    doc.add_heading("Valoración de Proyecto de Investigación", 1)
    doc.add_paragraph(f"Archivo: {nombre_archivo}")
    doc.add_paragraph(f"Resultado: {resultado}")
    doc.add_paragraph(f"Cumplimiento: {round(porcentaje, 2)}%")

    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje"

    for criterio, puntaje in scores.items():
        row = table.add_row().cells
        row[0].text = criterio
        row[1].text = str(puntaje)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


# ================= UI =================
st.set_page_config(page_title=APP_TITLE, page_icon="🧮", layout="wide")

st.title(APP_TITLE)
st.caption(APP_VERSION)

uploaded = st.file_uploader("Proyecto (PDF o DOCX)", type=["pdf", "docx"])

if uploaded is None:
    st.info("Esperando archivo…")
    st.stop()

raw = uploaded.read()

# Lectura básica del archivo solo para validar que abre
texto = ""
if uploaded.name.lower().endswith(".pdf"):
    if pdfplumber is None:
        st.error("Falta instalar pdfplumber.")
        st.stop()
    texto = parse_pdf(raw)
elif uploaded.name.lower().endswith(".docx"):
    if DocxDocument is None:
        st.error("Falta instalar python-docx.")
        st.stop()
    texto = parse_docx(raw)

if texto.strip():
    st.success("Archivo leído correctamente.")
else:
    st.warning("El archivo se cargó, pero no se pudo extraer texto visible.")

st.subheader("Evaluación manual")

scores = {}
total_max = sum(CRITERIOS.values())

cols = st.columns(2)
i = 0
for criterio, peso in CRITERIOS.items():
    with cols[i % 2]:
        st.markdown(f"**{criterio}** (máx. {peso})")
        val = st.slider(
            f"Puntaje - {criterio}",
            min_value=0,
            max_value=peso,
            value=peso,
            key=f"slider_{i}"
        )
        scores[criterio] = val
        st.divider()
    i += 1

total = sum(scores.values())
porcentaje = (total / total_max) * 100 if total_max else 0
resultado = categoria(porcentaje)

st.markdown(f"### Resultado: **{resultado}** — Cumplimiento **{round(porcentaje, 2)}%**")

c1, c2 = st.columns(2)

with c1:
    xls = make_excel(scores, porcentaje, resultado, uploaded.name)
    st.download_button(
        "⬇️ Descargar resultados.xlsx",
        data=xls,
        file_name="valoracion_proyecto.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with c2:
    docx_bytes = make_word(scores, porcentaje, resultado, uploaded.name)
    st.download_button(
        "⬇️ Descargar dictamen.docx",
        data=docx_bytes,
        file_name="dictamen_proyecto.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
