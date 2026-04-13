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
APP_VERSION = "v3.1 equilibrado"

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


def categoria(p):
    if p >= 70:
        return "Aprobado"
    elif p >= 50:
        return "Aprobado con observaciones"
    elif p >= 30:
        return "Requiere reformulación"
    return "No aprobado"


def parse_pdf(file_bytes):
    if pdfplumber is None:
        return ""
    partes = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            partes.append(page.extract_text() or "")
    return "\n".join(partes)


def parse_docx(file_bytes):
    if DocxDocument is None:
        return ""
    doc = DocxDocument(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs)


def make_excel(scores, porcentaje, resultado, nombre):
    filas = []
    for c, v in scores.items():
        filas.append({"Criterio": c, "Puntaje": v})

    df = pd.DataFrame(filas)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

        resumen = pd.DataFrame([{
            "Archivo": nombre,
            "Resultado": resultado,
            "Porcentaje": round(porcentaje, 2),
            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M")
        }])
        resumen.to_excel(writer, sheet_name="Resumen", index=False)

    return output.getvalue()


def make_word(scores, porcentaje, resultado, nombre):
    if Document is None:
        return b""

    doc = Document()
    doc.add_heading("Valoración de Proyecto de Investigación", 1)
    doc.add_paragraph(f"Archivo: {nombre}")
    doc.add_paragraph(f"Resultado: {resultado}")
    doc.add_paragraph(f"Cumplimiento: {round(porcentaje, 2)}%")

    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje"

    for c, v in scores.items():
        row = table.add_row().cells
        row[0].text = c
        row[1].text = str(v)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


# ================= UI =================

st.set_page_config(page_title=APP_TITLE, layout="wide")

st.title(APP_TITLE)
st.caption(APP_VERSION)

archivo = st.file_uploader("Subir proyecto (PDF o DOCX)", type=["pdf", "docx"])

if archivo is None:
    st.info("Esperando archivo…")
    st.stop()

raw = archivo.read()

texto = ""
if archivo.name.endswith(".pdf"):
    texto = parse_pdf(raw)
else:
    texto = parse_docx(raw)

st.success("Archivo cargado correctamente")

st.subheader("Evaluación")

scores = {}
total_max = sum(CRITERIOS.values())

cols = st.columns(2)
i = 0

for criterio, peso in CRITERIOS.items():
    with cols[i % 2]:
        st.markdown(f"**{criterio}** (máx {peso})")

        # 🔥 valor inicial equilibrado (75%)
        valor_inicial = max(1, round(peso * 0.75))

        val = st.slider(
            f"Puntaje {criterio}",
            0,
            peso,
            valor_inicial,
            key=f"s_{i}"
        )

        scores[criterio] = val
        st.divider()

    i += 1


total = sum(scores.values())
porcentaje = (total / total_max) * 100
resultado = categoria(porcentaje)

st.markdown(f"## Resultado: **{resultado}**")
st.markdown(f"### Cumplimiento: **{round(porcentaje,2)}%**")

c1, c2 = st.columns(2)

with c1:
    st.download_button(
        "⬇️ Descargar Excel",
        make_excel(scores, porcentaje, resultado, archivo.name),
        "resultado.xlsx"
    )

with c2:
    st.download_button(
        "⬇️ Descargar Word",
        make_word(scores, porcentaje, resultado, archivo.name),
        "resultado.docx"
    )
