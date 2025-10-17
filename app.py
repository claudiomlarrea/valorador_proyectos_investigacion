import io
import re
import yaml
import time
import base64
import hashlib
import textwrap
from datetime import datetime

import streamlit as st

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

APP_TITLE = "Valorador de Proyectos de Investigación"
APP_VERSION = "v1.0.0"

SECTION_KEYS = [
    "Resumen","Justificación","Relevancia","Planteamiento del problema",
    "Estado del arte","Marco teórico","Objetivo general","Objetivos específicos",
    "Pregunta de investigación","Metodología","Diseño y enfoque",
    "Datos y fuentes","Muestra y muestreo","Instrumentos",
    "Técnicas de análisis","Cronograma","Plan de actividades",
    "Resultados esperados","Impacto","Plan de difusión",
    "Gestión de riesgos","Ética","Presupuesto","Bibliografía","Referencias"
]

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
  "Bibliografía actualizada": {"peso": 8, "pistas": ["Bibliografía","Referencias","2020","2021","2022","2023","2024","2025"]}
}

THRESHOLDS = {
    "Aprobado": (60, 100),
    "Aprobado con observaciones": (50, 60),
    "No aprobado": (0, 50)
}

def load_criteria(file):
    try:
        cfg = yaml.safe_load(file)
        return cfg
    except Exception as e:
        st.error(f"No se pudo leer criteria.yaml: {e}")
        return None

def save_bytes_as(name, content):
    st.download_button("Descargar " + name, content, file_name=name)

def sanitize_text(txt):
    import re
    txt = re.sub(r'\r', ' ', txt)
    txt = re.sub(r'[ \t]+', ' ', txt)
    return txt

def parse_docx(file_bytes):
    if DocxDocument is None:
        return ""
    bio = io.BytesIO(file_bytes)
    doc = DocxDocument(bio)
    paras = [p.text for p in doc.paragraphs]
    return "\n".join(paras)

def parse_pdf(file_bytes):
    if pdfplumber is None:
        return ""
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
    return "\n".join(text_parts)

def heuristic_sections(text):
    found = {}
    for key in SECTION_KEYS:
        pattern = rf"(?im)^\s*{re.escape(key)}s*:?\s*$|^\s*{re.escape(key)}\s*$"
        matches = list(re.finditer(pattern, text))
        if matches:
            start = matches[0].end()
            next_positions = []
            for other in SECTION_KEYS:
                if other == key:
                    continue
                pat2 = rf"(?im)^\s*{re.escape(other)}s*:?\s*$|^\s*{re.escape(other)}\s*$"
                m2 = re.search(pat2, text[start:])
                if m2:
                    next_positions.append(start + m2.start())
            end = min(next_positions) if next_positions else len(text)
            found[key] = text[start:end].strip()
    if not found:
        import re as _re
        chunks = _re.split(r'\n{2,}', text)
        found = {f"Sección {i+1}": c for i, c in enumerate(chunks[:12])}
    return found

def compute_auto_hints(section_map, criteria_cfg):
    import re as _re
    hints = {}
    for criterio, meta in criteria_cfg.items():
        pistas = meta.get("pistas", [])
        evidence = []
        for p in pistas:
            for k, v in section_map.items():
                if _re.search(p, k, _re.IGNORECASE) or _re.search(p, v or "", _re.IGNORECASE):
                    snippet = (v or "")[:400].replace("\n", " ")
                    evidence.append(f"[{k}] {snippet}")
                    break
        hints[criterio] = evidence[:2]
    return hints

def score_block(criteria_cfg):
    total_peso = sum(int(v.get("peso", 0)) for v in criteria_cfg.values())
    st.write(f"**Puntaje total posible:** {total_peso} puntos")
    st.caption("Asigne puntajes por criterio (0–peso). El valor sugerido es sólo guía.")
    puntajes = {}
    for criterio, meta in criteria_cfg.items():
        peso = int(meta.get("peso", 0))
        with st.expander(f"{criterio} (peso {peso})", expanded=False):
            sugerido = meta.get("sugerido", None)
            if sugerido is None:
                sugerido = int(round(peso*0.7))
            val = st.slider("Puntaje", 0, peso, sugerido, key=f"score_{criterio}")
            evid = meta.get("evidencia", [])
            if evid:
                st.caption("Evidencia sugerida:")
                for e in evid:
                    st.code(e, language="markdown")
            obs = st.text_area("Observaciones", key=f"obs_{criterio}", placeholder="Notas, fortalezas, debilidades, recomendaciones…")
            puntajes[criterio] = {"asignado": val, "peso": peso, "observaciones": obs}
    obtenido = sum(v["asignado"] for v in puntajes.values())
    porcentaje = (obtenido / total_peso) * 100 if total_peso else 0.0
    return puntajes, obtenido, porcentaje, total_peso

def categorize(porcentaje):
    for label, (lo, hi) in THRESHOLDS.items():
        if lo <= porcentaje < hi or (label=="Aprobado" and porcentaje==hi):
            return label
    return "No clasificado"

def make_report(metadata, section_map, criteria_cfg, puntajes, obtenido, porcentaje, total_peso):
    lines = []
    lines.append(f"# Informe de Valoración — {APP_TITLE} {APP_VERSION}")
    lines.append(f"- Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"- Archivo: {metadata.get('filename','')}")
    lines.append(f"- Hash: `{metadata.get('sha256','')}`")
    lines.append("")
    lines.append("## Resumen ejecutivo")
    lines.append(f"- Puntaje total: **{obtenido}/{total_peso}** ({porcentaje:.1f}%)")
    lines.append(f"- Resultado: **{categorize(porcentaje)}**")
    lines.append("")
    lines.append("## Detalle por criterios")
    for criterio, meta in criteria_cfg.items():
        peso = int(meta.get("peso", 0))
        asignado = puntajes[criterio]["asignado"]
        obs = puntajes[criterio]["observaciones"]
        lines.append(f"### {criterio} — {asignado}/{peso}")
        if obs:
            lines.append(f"**Observaciones:** {obs}")
        if meta.get("evidencia"):
            lines.append("**Evidencia sugerida:**")
            for e in meta["evidencia"]:
                lines.append(f"- {e}")
        lines.append("")
    lines.append("## Secciones detectadas (extractos)")
    for k, v in section_map.items():
        if v:
            snippet = v.strip()[:1200]
            lines.append(f"### {k}\n\n{snippet}\n")
    return "\n".join(lines)

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧮", layout="wide")
    st.title(APP_TITLE)
    st.caption("Subí un proyecto (PDF o Word), evaluá con criterios configurables y exportá un informe.")

    with st.sidebar:
        st.markdown(f"**Versión:** {APP_VERSION}")
        st.markdown("### Umbrales")
        st.write("≥60% → **Aprobado**\n\n50–60% → **Aprobado con observaciones**\n\n<50% → **No aprobado**")
        st.divider()
        st.markdown("### Configuración de criterios")
        cfg_file = st.file_uploader("Opcional: cargar criteria.yaml", type=["yaml","yml"])
        if cfg_file:
            criteria_cfg = load_criteria(cfg_file)
        else:
            criteria_cfg = DEFAULT_CRITERIA
        with st.expander("Editar pesos (rápido)"):
            for k in list(criteria_cfg.keys()):
                peso = int(criteria_cfg[k].get("peso", 0))
                criteria_cfg[k]["peso"] = st.number_input(k, min_value=0, max_value=20, value=peso, key=f"peso_{k}")

    up = st.file_uploader("Proyecto en PDF o Word", type=["pdf","docx","doc"])
    if up is None:
        st.info("Esperando archivo…")
        return

    raw = up.read()
    sha = hashlib.sha256(raw).hexdigest()
    text = ""
    name = up.name.lower()

    if name.endswith(".pdf"):
        if pdfplumber is None:
            st.error("Falta dependencia: pdfplumber")
            return
        text = parse_pdf(raw)
    elif name.endswith(".docx"):
        if DocxDocument is None:
            st.error("Falta dependencia: python-docx")
            return
        text = parse_docx(raw)
    elif name.endswith(".doc"):
        try:
            text = raw.decode("latin-1", errors="ignore")
        except Exception:
            st.warning("Formato .doc antiguo. Convierta a PDF o DOCX para mejor extracción.")
            text = ""
    else:
        st.error("Formato no soportado.")
        return

    text = sanitize_text(text)
    if not text.strip():
        st.error("No se pudo extraer texto. Convierta el archivo a PDF o DOCX y vuelva a intentar.")
        return

    sections = heuristic_sections(text)

    st.subheader("Secciones detectadas")
    cols = st.columns(2)
    keys = list(sections.keys())
    for i, k in enumerate(keys):
        with cols[i % 2]:
            st.markdown(f"**{k}**")
            st.code(textwrap.shorten(sections[k], width=600, placeholder='…'))

    hints = compute_auto_hints(sections, criteria_cfg)
    for c, meta in criteria_cfg.items():
        meta["evidencia"] = hints.get(c, [])

    st.subheader("Valoración")
    puntajes, obtenido, porcentaje, total = score_block(criteria_cfg)

    c1, c2, c3 = st.columns(3)
    c1.metric("Puntaje", f"{obtenido}/{total}")
    c2.metric("Porcentaje", f"{porcentaje:.1f}%")
    c3.metric("Resultado", categorize(porcentaje))

    if st.button("Generar informe (Markdown)"):
        meta = {"filename": up.name, "sha256": sha}
        md = make_report(meta, sections, criteria_cfg, puntajes, obtenido, porcentaje, total)
        b = md.encode("utf-8")
        save_bytes_as(f"valoracion_{pathlib.Path(up.name).stem}.md", b)

    st.divider()
    st.caption("Consejo: los criterios y pesos son editables. Guardá tu criteria.yaml para reutilizar.")
    st.code(yaml.safe_dump(DEFAULT_CRITERIA, sort_keys=False), language="yaml")

if __name__ == "__main__":
    main()