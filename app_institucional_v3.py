import io
import re
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
APP_VERSION = "v3.3 diferenciación real por contenido"

CRITERIOS = {
    "Pertinencia y relevancia": {
        "peso": 10,
        "pistas": ["justificación", "relevancia", "problema", "fundamentación", "necesidad"]
    },
    "Claridad del problema y objetivos": {
        "peso": 10,
        "pistas": ["objetivo general", "objetivos específicos", "pregunta de investigación", "problema", "hipótesis"]
    },
    "Originalidad / aporte": {
        "peso": 8,
        "pistas": ["estado del arte", "marco teórico", "antecedentes", "novedad", "aporte", "vacancia"]
    },
    "Solidez metodológica": {
        "peso": 14,
        "pistas": ["metodología", "diseño", "enfoque", "técnicas", "análisis de datos", "método"]
    },
    "Calidad de datos / muestra": {
        "peso": 10,
        "pistas": ["muestra", "muestreo", "población", "instrumento", "datos", "recolección"]
    },
    "Factibilidad y cronograma": {
        "peso": 8,
        "pistas": ["cronograma", "plan de actividades", "factibilidad", "recursos", "viabilidad", "etapas"]
    },
    "Consideraciones éticas": {
        "peso": 6,
        "pistas": ["ética", "consentimiento", "confidencialidad", "comité de ética", "resguardo de datos"]
    },
    "Impacto esperado": {
        "peso": 8,
        "pistas": ["impacto", "resultados esperados", "beneficios", "relevancia social", "aportes"]
    },
    "Plan de difusión / transferencia": {
        "peso": 6,
        "pistas": ["difusión", "transferencia", "publicaciones", "divulgación", "congreso", "artículo"]
    },
    "Presupuesto y sostenibilidad": {
        "peso": 6,
        "pistas": ["presupuesto", "financiamiento", "costos", "recursos", "gastos", "sostenibilidad"]
    },
    "Alineación institucional y normativa": {
        "peso": 6,
        "pistas": ["institucional", "normativa", "lineamientos", "universidad", "facultad", "plan estratégico"]
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


def contar_ocurrencias(texto, termino):
    return texto.lower().count(termino.lower())


def extraer_evidencia(texto, pistas, max_items=2):
    texto_low = texto.lower()
    resultados = []

    for pista in pistas:
        idx = texto_low.find(pista.lower())
        if idx != -1:
            inicio = max(0, idx - 80)
            fin = min(len(texto), idx + 180)
            frag = texto[inicio:fin].replace("\n", " ").strip()
            if frag not in resultados:
                resultados.append(frag)
        if len(resultados) >= max_items:
            break

    return resultados


def score_criterio(texto, criterio, meta):
    """
    Genera un puntaje inicial variable y mucho más fino.
    Nunca deja todos iguales salvo que los proyectos sean realmente casi idénticos.
    """
    peso = meta["peso"]
    pistas = meta["pistas"]
    texto_low = texto.lower()

    # 1) cuántas pistas distintas aparecen
    hits_distintos = sum(1 for p in pistas if p.lower() in texto_low)

    # 2) cuántas ocurrencias totales hay
    ocurrencias_totales = sum(contar_ocurrencias(texto, p) for p in pistas)

    # 3) densidad del texto (cantidad total de palabras)
    n_palabras = max(1, len(texto.split()))

    # 4) bonus por desarrollo textual general del proyecto
    if n_palabras >= 5000:
        bonus_longitud = 0.12
    elif n_palabras >= 2500:
        bonus_longitud = 0.08
    elif n_palabras >= 1200:
        bonus_longitud = 0.05
    else:
        bonus_longitud = 0.02

    # 5) score base por criterio
    proporcion_hits = hits_distintos / max(1, len(pistas))
    factor_ocurrencias = min(1.0, ocurrencias_totales / max(2, len(pistas) * 2))

    # base entre 35% y 90% del peso
    score_relativo = 0.15 + (proporcion_hits * 0.35) + (factor_ocurrencias * 0.30) + bonus_longitud

    # Ajuste especial para bibliografía actualizada
    if criterio == "Bibliografía actualizada":
        years = re.findall(r"\b(2021|2022|2023|2024|2025|2026)\b", texto)
        years_unicos = len(set(years))
        score_relativo += min(0.15, years_unicos * 0.03)

    # Ajuste especial para presupuesto: detectar números o moneda
    if criterio == "Presupuesto y sostenibilidad":
        if re.search(r"(\$|usd|ars|presupuesto|costos|gastos|financiamiento)", texto_low):
            score_relativo += 0.08
        if re.search(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?\b", texto):
            score_relativo += 0.05

    # Ajuste especial para metodología
    if criterio == "Solidez metodológica":
        if re.search(r"(cuantitativ|cualitativ|mixto|estadístic|entrevista|encuesta|análisis)", texto_low):
            score_relativo += 0.08

    # Ajuste especial para muestra/datos
    if criterio == "Calidad de datos / muestra":
        if re.search(r"(n=|muestra|muestreo|población|casos|participantes|instrumento)", texto_low):
            score_relativo += 0.08

    # Limitar entre 35% y 95%
    score_relativo = max(0.35, min(0.95, score_relativo))

    valor = round(peso * score_relativo)

    # que nunca sea 0 si el criterio existe, pero tampoco siempre igual
    valor = max(1, min(peso, valor))

    return valor, hits_distintos, ocurrencias_totales


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
st.markdown("""
<style>

/* Fondo general */
.stApp {
    background-color: #E6E6E6;
}

/* Título principal */
h1 {
    color: #064a3f !important;
    font-weight: 700;
}

/* Subtítulos */
h2, h3, h4 {
    color: #064a3f !important;
}

/* Texto */
p, label, span {
    color: black !important;
}

/* Caja de carga */
[data-testid="stFileUploader"] {
    background-color: white;
    border-radius: 10px;
    padding: 15px;
}

/* Alertas (éxito, info, warning) */
[data-testid="stAlert"] {
    border-radius: 10px;
}

/* Botones */
.stButton button {
    background-color: #064a3f;
    color: white;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}

.stButton button:hover {
    background-color: #0B6B5D;
}

/* Sliders */
[data-baseweb="slider"] {
    color: #064a3f;
}

</style>
""", unsafe_allow_html=True)

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

        valor_inicial, hits_distintos, ocurrencias_totales = score_criterio(texto, criterio, meta)
        evidencias = extraer_evidencia(texto, meta["pistas"])

        st.markdown(f"**{criterio}** (máx {peso})")
        st.caption(f"Pistas detectadas: {hits_distintos} | Ocurrencias: {ocurrencias_totales}")

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
