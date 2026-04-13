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
APP_VERSION = "v3.2 equilibrado con lectura automática"

CRITERIOS = {
    "Pertinencia y relevancia": {
        "peso": 10,
        "pistas": ["justificación", "relevancia", "problema", "fundamentación"]
    },
    "Claridad del problema y objetivos": {
        "peso": 10,
        "pistas": ["objetivo general", "objetivos específicos", "pregunta de investigación", "problema"]
    },
    "Originalidad / aporte": {
        "peso": 8,
        "pistas": ["estado del arte", "marco teórico", "antecedentes", "novedad", "aporte"]
    },
    "Solidez metodológica": {
        "peso": 14,
        "pistas": ["metodología", "diseño", "enfoque", "técnicas", "análisis de datos"]
    },
    "Calidad de datos / muestra": {
        "peso": 10,
        "pistas": ["muestra", "muestreo", "población", "instrumento", "datos"]
    },
    "Factibilidad y cronograma": {
        "peso": 8,
        "pistas": ["cronograma", "plan de actividades", "factibilidad", "recursos"]
    },
    "Consideraciones éticas": {
        "peso": 6,
        "pistas": ["ética", "consentimiento", "confidencialidad", "comité de ética"]
    },
    "Impacto esperado": {
        "peso": 8,
        "pistas": ["impacto", "resultados esperados", "beneficios", "relevancia social"]
    },
    "Plan de difusión / transferencia": {
        "peso": 6,
        "pistas": ["difusión", "transferencia", "publicaciones", "divulgación", "congreso"]
    },
    "Presupuesto y sostenibilidad": {
        "peso": 6,
        "pistas": ["presupuesto", "financiamiento", "costos", "recursos", "gastos"]
    },
    "Alineación institucional y normativa": {
        "peso": 6,
        "pistas": ["institucional", "normativa", "lineamientos", "universidad", "facultad"]
    },
    "Bibliografía actualizada": {
        "peso": 8,
        "pistas": ["bibliografía", "referencias", "2021", "2022", "2023", "2024", "2025", "2026"]
    },
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


def detectar_nivel_evidencia(texto, pistas):
    """
    Devuelve 0, 1 o 2:
    0 = sin evidencia
    1 = evidencia parcial
    2 = evidencia suficiente
    """
    texto_low = texto.lower()
    hits = 0
    for pista in pistas:
        if pista.lower() in texto_low:
            hits += 1

    if hits >= 3:
        return 2
    elif hits >= 1:
        return 1
    return 0


def puntaje_inicial(peso, nivel):
    """
    Nivel 2: 85%
    Nivel 1: 70%
    Nivel 0: 50%
    """
    if nivel == 2:
        return max(1, round(peso * 0.85))
    elif nivel == 1:
        return max(1, round(peso * 0.70))
    else:
        return max(1, round(peso * 0.50))


def extraer_evidencia(texto, pistas, max_items=2):
    texto_low = texto.lower()
    resultados = []

    for pista in pistas:
        idx = texto_low.find(pista.lower())
        if idx != -1:
            inicio = max(0, idx - 80)
            fin = min(len(texto), idx + 160)
            frag = texto[inicio:fin].replace("\n", " ").strip()
            if frag not in resultados:
                resultados.append(frag)
        if len(resultados) >= max_items:
            break

    return resultados


def make_excel(scores, porcentaje, resultado, nombre):
    filas = []
    for c, v in scores.items():
        filas.append({"Criterio": c, "Puntaje": v})

    df = pd.DataFrame(filas)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")

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
if archivo.name.lower().endswith(".pdf"):
    texto = parse_pdf(raw)
else:
    texto = parse_docx(raw)

if texto.strip():
    st.success("Archivo cargado correctamente")
else:
    st.warning("Se cargó el archivo, pero no se extrajo texto visible.")

st.subheader("Evaluación")

scores = {}
total_max = sum(meta["peso"] for meta in CRITERIOS.values())

cols = st.columns(2)
i = 0

for criterio, meta in CRITERIOS.items():
    with cols[i % 2]:
        peso = meta["peso"]
        pistas = meta["pistas"]

        nivel = detectar_nivel_evidencia(texto, pistas)
        valor_inicial = puntaje_inicial(peso, nivel)
        evidencias = extraer_evidencia(texto, pistas)

        st.markdown(f"**{criterio}** (máx {peso})")

        if nivel == 2:
            st.success("Evidencia suficiente detectada.")
        elif nivel == 1:
            st.warning("Evidencia parcial detectada.")
        else:
            st.info("No se detectó evidencia clara. Revisar manualmente.")

        val = st.slider(
            f"Puntaje {criterio}",
            0,
            peso,
            valor_inicial,
            key=f"s_{i}"
        )

        if evidencias:
            with st.expander("Evidencia sugerida"):
                for ev in evidencias:
                    st.write(ev)

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
