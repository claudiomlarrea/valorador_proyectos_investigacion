
import io, yaml, textwrap, hashlib
from datetime import datetime


# --- FIX AUTO: quitar SOLO la sección "Extracto de evidencia del documento" del Word ---
def _delete_paragraph(p):
    try:
        p._element.getparent().remove(p._element)
        p._p = p._element = None
    except Exception:
        pass

def remove_section_by_heading(doc, heading_text="Extracto de evidencia del documento"):
    """Elimina el heading 'Extracto de evidencia del documento' y su contenido hasta el siguiente heading."""
    try:
        paras = list(doc.paragraphs)
        i = 0
        while i < len(paras):
            p = paras[i]
            txt = (p.text or "").strip()
            if txt.lower() == (heading_text or "").lower():
                # borrar el heading
                _delete_paragraph(p)
                # borrar párrafos hasta el próximo heading
                # se considera heading si el nombre de estilo arranca con 'Heading' o 'Título' (Word en español)
                j = i  # ya avanzamos al siguiente por cómo funciona la lista original
                # debemos refrescar la referencia a doc.paragraphs porque se van borrando
                while j < len(doc.paragraphs):
                    pj = doc.paragraphs[j]
                    style_name = (getattr(getattr(pj, "style", None), "name", "") or "").lower()
                    if style_name.startswith("heading") or style_name.startswith("título"):
                        break
                    _delete_paragraph(pj)
                break
            i += 1
    except Exception:
        # Silencioso: nunca rompe la exportación
        pass

import streamlit as st

# Parsing libs
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except Exception:
    Document = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

import pandas as pd

APP_TITLE = "UCCuyo · Valorador de Proyectos de Investigación"
APP_VERSION = "v1.1 – Excel & Word"

# --- Criterios y pesos por defecto (ex ante) ---
DEFAULT_CRITERIA = {
  "Pertinencia y relevancia": {"peso": 10, "pistas": ["Justificación","Relevancia","Problema"]},
  "Claridad del problema y objetivos": {"peso": 10, "pistas": ["Planteamiento del problema","Objetivo general","Objetivos específicos","Pregunta de investigación"]},
  "Originalidad / aporte": {"peso": 8, "pistas": ["Estado del arte","Marco teórico","Novedad"]},
  "Solidez metodológica": {"peso": 14, "pistas": ["Metodología","Diseño y enfoque","Técnicas de análisis"]},
  "Calidad de datos / muestra": {"peso": 10, "pistas": ["Datos","Muestra","Muestreo","Instrumentos"]},
  "Factibilidad y cronograma": {"peso": 8, "pistas": ["Cronograma","Plan de actividades","Recursos"]},
  "Consideraciones éticas": {"peso": 6, "pistas": ["Ética","Consentimiento","Privacidad"]},
  "Impacto esperado": {"peso": 8, "pistas": ["Resultados esperados","Impacto","Relevancia social"]},
  "Plan de difusión / transferencia": {"peso": 6, "pistas": ["Plan de difusión","Transferencia","Publicaciones"]},
  "Presupuesto y sostenibilidad": {"peso": 6, "pistas": ["Presupuesto","Recursos","Financiamiento"]},
  "Alineación institucional y normativa": {"peso": 6, "pistas": ["Institucional","Lineamientos","Normativa"]},
  "Bibliografía actualizada": {"peso": 8, "pistas": ["Bibliografía","Referencias","2021","2022","2023","2024","2025"]}
}

THRESHOLDS = {
    "Aprobado": (60, 1000),  # 60–100
    "Aprobado con observaciones": (50, 60),
    "No aprobado": (0, 50)
}

def categorize(porcentaje: float) -> str:
    for label, (lo, hi) in THRESHOLDS.items():
        if lo <= porcentaje < hi or (label == "Aprobado" and porcentaje == hi):
            return label
    return "No clasificado"

def parse_docx(file_bytes: bytes) -> str:
    if DocxDocument is None:
        return ""
    bio = io.BytesIO(file_bytes)
    doc = DocxDocument(bio)
    return "\n".join(p.text for p in doc.paragraphs)

def parse_pdf(file_bytes: bytes) -> str:
    if pdfplumber is None:
        return ""
    parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for pg in pdf.pages:
            parts.append(pg.extract_text() or "")
    return "\n".join(parts)

def compute_auto_hints(text: str, criteria_cfg: dict) -> dict:
    low = text.lower()
    hints = {}
    for crit, meta in criteria_cfg.items():
        pistas = meta.get("pistas", [])
        ev = []
        for p in pistas:
            if p.lower() in low:
                # save a short snippet around the first occurrence
                idx = low.find(p.lower())
                a = max(0, idx-120)
                b = min(len(text), idx+200)
                ev.append(text[a:b].replace("\n", " "))
        hints[crit] = ev[:2]
    return hints

def score_ui(criteria_cfg: dict):
    total_peso = sum(int(v.get("peso", 0)) for v in criteria_cfg.values())
    st.caption(f"Puntaje total posible: {total_peso} puntos (suma de pesos)")
    puntajes = {}
    cols = st.columns(2)
    i = 0
    for crit, meta in criteria_cfg.items():
        with cols[i % 2]:
            peso = int(meta.get("peso", 0))
            st.markdown(f"**{crit}**  (peso {peso})")
            val = st.slider("Puntaje asignado", 0, peso, int(round(peso*0.7)), key=f"score_{crit}")
            obs = st.text_area("Observaciones", key=f"obs_{crit}", placeholder="Notas, fortalezas, debilidades, recomendaciones…")
            if meta.get("evidencia"):
                with st.expander("Evidencia sugerida (auto)", expanded=False):
                    for e in meta["evidencia"]:
                        st.code(e, language="markdown")
            puntajes[crit] = {"asignado": val, "peso": peso, "observaciones": obs}
            st.divider()
        i += 1
    obtenido = sum(v["asignado"] for v in puntajes.values())
    porcentaje = (obtenido / total_peso) * 100 if total_peso else 0.0
    return puntajes, obtenido, porcentaje, total_peso

def make_excel(criteria_cfg, puntajes: dict, porcentaje: float, result: str, nombre_archivo: str) -> bytes:
    weights_df = []
    for crit, meta in criteria_cfg.items():
        peso = int(meta.get("peso", 0))
        asignado = puntajes[crit]["asignado"]
        aporte = round((asignado / peso) * peso if peso>0 else 0, 2)  # igual al asignado, se deja explícito
        weights_df.append({
            "Criterio": crit,
            "Peso": peso,
            "Puntaje asignado": asignado,
            "Aporte (%)": round((asignado / sum(int(v.get('peso',0)) for v in criteria_cfg.values()))*100, 2),
            "Observaciones": puntajes[crit]["observaciones"]
        })
    df = pd.DataFrame(weights_df)

    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Resultados")
            resumen = pd.DataFrame([{
                "Archivo": nombre_archivo,
                "Resultado": result,
                "Porcentaje total": round(porcentaje,2),
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M")
            }])
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
        return output.getvalue()

def make_word(criteria_cfg, puntajes: dict, porcentaje: float, result: str, nombre_archivo: str, extracto: str) -> bytes:
    if Document is None:
        return b""
    doc = Document()
    styles = doc.styles['Normal']
    styles.font.name = 'Times New Roman'
    styles.font.size = Pt(11)

    # Encabezado institucional simple
    h = doc.add_paragraph("Universidad Católica de Cuyo – Secretaría de Investigación")
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_format = h.runs[0].font
    h_format.size = Pt(12)

    doc.add_heading("Valoración de Proyecto de Investigación", level=1)
    meta = doc.add_paragraph()
    meta.add_run(f"Archivo: ").bold = True
    meta.add_run(nombre_archivo + "   ")
    meta.add_run("Fecha: ").bold = True
    meta.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))

    doc.add_paragraph("")
    doc.add_paragraph(f"Dictamen: {result} — Cumplimiento: {round(porcentaje,2)}%")

    doc.add_heading("Resultados por criterio", level=2)
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Peso"
    hdr[2].text = "Puntaje asignado"
    hdr[3].text = "Observaciones"

    for crit, meta in criteria_cfg.items():
        row = table.add_row().cells
        row[0].text = crit
        row[1].text = str(int(meta.get("peso",0)))
        row[2].text = str(puntajes[crit]["asignado"])
        row[3].text = puntajes[crit]["observaciones"] or ""

    # Fortalezas y mejoras
    fortalezas = [c for c in criteria_cfg if puntajes[c]["asignado"] >= int(criteria_cfg[c]["peso"]*0.75)]
    mejoras = [c for c in criteria_cfg if puntajes[c]["asignado"] <= int(criteria_cfg[c]["peso"]*0.25)]

    doc.add_paragraph("")
    doc.add_heading("Síntesis", level=2)
    doc.add_paragraph("Fortalezas: " + (", ".join(fortalezas) if fortalezas else "No se destacan fortalezas específicas."))
    doc.add_paragraph("Aspectos a mejorar: " + (", ".join(mejoras) if mejoras else "No se identifican aspectos críticos."))

    doc.add_paragraph("")
    doc.add_heading("Extracto de evidencia del documento", level=2)
    doc.add_paragraph(extracto[:2000])

    with io.BytesIO() as buffer:
remove_section_by_heading(doc)
        doc.save(buffer)
        return buffer.getvalue()

# ================== UI ==================
st.set_page_config(page_title=APP_TITLE, page_icon="🧮", layout="wide")
st.title(APP_TITLE)
st.caption("Cargá un proyecto (PDF/DOCX). Asigná puntajes por criterio y exportá Excel y Word con el dictamen.")

uploaded = st.file_uploader("Proyecto (PDF o DOCX)", type=["pdf", "docx"])

if uploaded is None:
    st.info("Esperando archivo…")
    st.stop()

raw = uploaded.read()
sha = hashlib.sha256(raw).hexdigest()
text = ""

if uploaded.name.lower().endswith(".pdf"):
    if pdfplumber is None:
        st.error("Falta dependencia: pdfplumber")
        st.stop()
    text = parse_pdf(raw)
else:
    if DocxDocument is None:
        st.error("Falta dependencia: python-docx")
        st.stop()
    text = parse_docx(raw)

text_low = text.lower()

# Criterios
criteria_cfg = DEFAULT_CRITERIA.copy()
# Sugerencias automáticas de evidencia
hints = compute_auto_hints(text, criteria_cfg)
for c in criteria_cfg:
    criteria_cfg[c]["evidencia"] = hints.get(c, [])

# UI de puntajes
puntajes, obtenido, porcentaje, total_peso = score_ui(criteria_cfg)

# Resultado
resultado = categorize(porcentaje)
st.markdown(f"### Resultado: **{resultado}** — Cumplimiento **{round(porcentaje,2)}%**")

col1, col2 = st.columns(2)
with col1:
    if st.button("⬇️ Exportar Excel"):
        xls = make_excel(criteria_cfg, puntajes, porcentaje, resultado, uploaded.name)
        st.download_button("Descargar resultados.xlsx", data=xls, file_name="valoracion_proyecto.xlsx")
with col2:
    if st.button("⬇️ Exportar Word"):
        docx_bytes = make_word(criteria_cfg, puntajes, porcentaje, resultado, uploaded.name, text[:4000])
        st.download_button("Descargar dictamen.docx", data=docx_bytes, file_name="dictamen_proyecto.docx")

st.divider()
st.caption("Nota: Este valorador es ex ante. No se genera informe Markdown ni se muestran bloques de depuración.")
