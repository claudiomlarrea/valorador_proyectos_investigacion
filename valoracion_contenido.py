"""
Valoración ex ante por contenido propio (Anexo IV + plantillas UCCuyo reales).
- No penaliza: solo otorga puntos por evidencia.
- Excluye consignas de plantilla / instructivo.
- Si la plantilla no coincide (p. ej. informe de cátedra), usa respaldo sobre texto propio global.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

# Encabezados reales (convocatoria interna + variaciones)
MAIN_SECTION_HEADERS: List[Tuple[str, str]] = [
    ("identificacion", r"(?i)^\s*(?:\d+\.?\s*)?identificaci[oó]n(?:\s+del\s+proyecto)?\b"),
    ("director", r"(?i)^\s*(?:\d+\.?\s*)?datos\s+(?:del\s+)?director"),
    ("equipo", r"(?i)^\s*(?:\d+\.?\s*)?equipo\s+de\s+investigaci[oó]n\b"),
    ("resumen", r"(?i)^\s*(?:\d+\.?\s*)?resumen(?:\s+del\s+proyecto)?\b"),
    ("fundamentacion", r"(?i)^\s*(?:\d+\.?\s*)?fundamentaci[oó]n\b"),
    ("pertinencia", r"(?i)^\s*(?:\d+\.?\s*)?pertinencia\s+y\s+relevancia\b"),
    (
        "problema_objetivos",
        r"(?i)^\s*(?:\d+\.?\s*)?(?:(?:planteo\s+del\s+)?problema\s+y\s+objetivos|planteo\s+del\s+problema)\b",
    ),
    (
        "originalidad",
        r"(?i)^\s*(?:\d+\.?\s*)?originalidad(?:\s+y\s+aporte(?:\s+al\s+conocimiento)?)?\b",
    ),
    (
        "marco_estado",
        r"(?i)^\s*(?:\d+\.?\s*)?(?:marco\s+te[oó]rico\s+y\s+estado\s+del\s+arte|estado\s+del\s+arte)\b",
    ),
    ("metodologia", r"(?i)^\s*(?:\d+\.?\s*)?metodolog[ií]a\b"),
    (
        "factibilidad",
        r"(?i)^\s*(?:\d+\.?\s*)?(?:factibilidad\s+y\s+cronograma|factibilidad)\b",
    ),
    ("etica", r"(?i)^\s*(?:\d+\.?\s*)?(?:consideraciones\s+)?[eé]tica"),
    (
        "impacto_difusion",
        r"(?i)^\s*(?:\d+\.?\s*)?(?:impacto\s+esperado\s+y\s+plan\s+de\s+difusi[oó]n|impacto\s+esperado)\b",
    ),
    (
        "presupuesto",
        r"(?i)^\s*(?:\d+\.?\s*)?(?:presupuesto\s*,\s*sostenibilidad\s+y\s+alineaci[oó]n|presupuesto(?:\s*,\s*sostenibilidad)?)\b",
    ),
    ("bibliografia", r"(?i)^\s*(?:\d+\.?\s*)?bibliograf[ií]a\b"),
]

# Informe final de cátedra (estructura distinta)
INFORME_CATEDRA_HEADERS: List[Tuple[str, str]] = [
    ("identificacion", r"(?i)^\s*secci[oó]n\s+1\b"),
    ("resumen", r"(?i)^\s*(?:secci[oó]n\s+2\b|resumen\s+ejecutivo)"),
    ("factibilidad", r"(?i)^\s*(?:secci[oó]n\s+3\.?\s*1\b|cronograma\s+y\s+objetivos)"),
    ("impacto_difusion", r"(?i)^\s*(?:secci[oó]n\s+3\.?\s*2\b|producci[oó]n\s+y\s+transferencia)"),
    ("metodologia", r"(?i)^\s*metodolog[ií]a\b"),
    ("presupuesto", r"(?i)^\s*presupuesto\b"),
]


def _compile_header_patterns(
    headers: List[Tuple[str, str]],
) -> List[Tuple[str, re.Pattern[str]]]:
    compiled: List[Tuple[str, re.Pattern[str]]] = []
    for key, pat in headers:
        try:
            compiled.append((key, re.compile(pat)))
        except re.error as exc:
            raise re.error(f"Patrón inválido para sección '{key}': {pat}") from exc
    return compiled


MAIN_SECTION_HEADER_RX = _compile_header_patterns(MAIN_SECTION_HEADERS)
INFORME_CATEDRA_HEADER_RX = _compile_header_patterns(INFORME_CATEDRA_HEADERS)

SECTION_WORD_RANGES: Dict[str, Tuple[int, int]] = {
    "resumen": (80, 350),
    "fundamentacion": (200, 700),
    "pertinencia": (100, 450),
    "problema_objetivos": (100, 450),
    "originalidad": (60, 250),
    "marco_estado": (300, 900),
    "metodologia": (400, 1000),
    "factibilidad": (100, 400),
    "impacto_difusion": (100, 400),
    "presupuesto": (80, 350),
    "bibliografia": (80, 400),
    "cuerpo": (400, 8000),
}

CRITERION_SECTIONS: Dict[str, List[str]] = {
    "Pertinencia y relevancia": ["fundamentacion", "pertinencia", "cuerpo"],
    "Claridad del problema y objetivos": ["resumen", "problema_objetivos", "cuerpo"],
    "Originalidad / aporte": ["originalidad", "marco_estado", "cuerpo"],
    "Solidez metodológica": ["metodologia", "problema_objetivos", "cuerpo"],
    "Calidad de datos / muestra": ["metodologia", "cuerpo"],
    "Factibilidad y cronograma": ["factibilidad", "cuerpo"],
    "Consideraciones éticas": ["etica", "metodologia", "cuerpo"],
    "Impacto esperado": ["resumen", "impacto_difusion", "cuerpo"],
    "Plan de difusión / transferencia": ["impacto_difusion", "cuerpo"],
    "Presupuesto y sostenibilidad": ["presupuesto", "factibilidad", "cuerpo"],
    "Alineación institucional y normativa": ["pertinencia", "presupuesto", "cuerpo"],
    "Bibliografía actualizada": ["marco_estado", "bibliografia", "cuerpo"],
}

CRITERION_CHECKS: Dict[str, List[str]] = {
    "Pertinencia y relevancia": [
        "justificación", "relevancia", "problema", "fundamentación", "pei", "plan estratégico", "líneas",
    ],
    "Claridad del problema y objetivos": [
        "objetivo general", "objetivos específicos", "objetivo", "pregunta", "problema", "hipótesis", "variables",
    ],
    "Originalidad / aporte": [
        "estado del arte", "marco teórico", "antecedentes", "novedad", "aporte", "vacío", "vacancia",
    ],
    "Solidez metodológica": [
        "metodología", "metodolog", "diseño", "enfoque", "cuantitativo", "cualitativo", "mixto",
        "técnicas", "análisis", "procedimiento",
    ],
    "Calidad de datos / muestra": [
        "muestra", "muestreo", "población", "instrumento", "participantes", "n=", "criterios de inclusión",
    ],
    "Factibilidad y cronograma": [
        "cronograma", "plan de actividades", "factibilidad", "recursos", "viabilidad", "etapas", "riesgos", "meses",
    ],
    "Consideraciones éticas": [
        "ética", "consentimiento", "confidencialidad", "comité", "resguardo",
    ],
    "Impacto esperado": [
        "impacto", "resultados esperados", "beneficios", "relevancia social",
    ],
    "Plan de difusión / transferencia": [
        "difusión", "transferencia", "publicaciones", "divulgación", "congreso", "artículo",
    ],
    "Presupuesto y sostenibilidad": [
        "presupuesto", "financiamiento", "costos", "gastos", "sostenibilidad", "fuente", "ars", "$",
    ],
    "Alineación institucional y normativa": [
        "institucional", "normativa", "lineamientos", "universidad", "facultad", "pei", "alineación",
    ],
    "Bibliografía actualizada": ["bibliografía", "referencias"],
}

CONSIGNA_LINE_PATTERNS = [
    r"(?i)^\s*debe\s+incluir",
    r"(?i)^\s*debe\s+demostrar",
    r"(?i)^\s*describir\b",
    r"(?i)^\s*indicar\b",
    r"(?i)^\s*explicar\b",
    r"(?i)^\s*presentar\b",
    r"(?i)^\s*mencionar\b",
    r"(?i)^\s*especificar\b",
    r"(?i)^\s*formular\b",
    r"(?i)^\s*evitar\s*:",
    r"(?i)^\s*regla\s+fundamental",
    r"(?i)^\s*notas\s+sobre",
    r"(?i)^\s*formatos\s+permitidos",
    r"(?i)^\s*condiciones\s+obligatorias",
    r"(?i)^\s*importante\s*:",
    r"(?i)extensi[oó]n\s+sugerida",
    r"(?i)^\s*en\s+esta\s+etapa\b",
    r"(?i)^\s*se\s+realizar[aá]\b",
    r"(?i)^\s*se\s+llevar[aá]\b",
    r"(?i)^\s*se\s+proceder[aá]\b",
]


@dataclass
class SectionSlice:
    key: str
    raw: str
    own: str
    word_count_own: int
    consigna_ratio: float


@dataclass
class CriterionScore:
    puntaje: int
    peso_max: int
    checks_ok: int
    checks_total: int
    word_own: int
    word_target_min: int
    word_target_max: int
    evidencias: List[str] = field(default_factory=list)
    notas: List[str] = field(default_factory=list)
    modo: str = ""


def _normalize(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _word_count(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text or "", flags=re.UNICODE))


def detect_doc_mode(text: str) -> str:
    head = (text or "")[:2500].lower()
    if "informe final" in head and ("cátedra" in head or "catedra" in head):
        return "informe_catedra"
    return "proyecto_ex_ante"


def _is_consigna_line(line: str) -> bool:
    s = line.strip()
    if len(s) < 8:
        return False
    for pat in CONSIGNA_LINE_PATTERNS:
        if re.search(pat, s):
            return True
    if re.match(r"(?i)^\s*•\s+", s) and _word_count(s) < 18:
        return True
    return False


def extract_own_content(block: str) -> Tuple[str, float]:
    if not block or not block.strip():
        return "", 1.0
    kept: List[str] = []
    consigna_lines = 0
    total = 0
    for line in block.split("\n"):
        line = line.strip()
        if not line:
            continue
        total += 1
        if _is_consigna_line(line):
            consigna_lines += 1
            continue
        kept.append(line)
    own = "\n".join(kept).strip()
    return own, consigna_lines / max(1, total)


def _line_is_main_header(
    line: str, patterns: List[Tuple[str, re.Pattern[str]]]
) -> Tuple[str, str] | None:
    s = line.strip()
    if not s or len(s) > 120:
        return None
    # Sub-etapas del cronograma: "1. Diseño del sistema (Meses 1–2)" sin palabra de sección principal
    if re.match(r"(?i)^\d+\.\s+", s):
        if not re.search(
            r"(?i)identificaci[oó]n|director|equipo|resumen|fundamentaci[oó]n|pertinencia|problema|"
            r"originalidad|marco|metodolog|factibilidad|impacto|presupuesto|bibliograf|secci[oó]n\s+\d",
            s,
        ):
            return None
    for key, rx in patterns:
        if rx.search(s):
            return key, s
    return None


def split_sections(full_text: str) -> Tuple[Dict[str, str], str]:
    text = _normalize(full_text)
    if not text:
        return {}, "proyecto_ex_ante"

    mode = detect_doc_mode(text)
    patterns = (
        INFORME_CATEDRA_HEADER_RX if mode == "informe_catedra" else MAIN_SECTION_HEADER_RX
    )

    lines = text.split("\n")
    markers: List[Tuple[int, str]] = []
    for i, line in enumerate(lines):
        hit = _line_is_main_header(line, patterns)
        if hit:
            markers.append((i, hit[0]))

    if not markers:
        return {"cuerpo": text}, mode

    sections: Dict[str, str] = {}
    for idx, (start_line, key) in enumerate(markers):
        end_line = markers[idx + 1][0] if idx + 1 < len(markers) else len(lines)
        block_lines = lines[start_line + 1 : end_line]
        block = "\n".join(block_lines).strip()
        if key in sections:
            sections[key] = (sections[key] + "\n\n" + block).strip()
        else:
            sections[key] = block

    return sections, mode


def build_section_slices(full_text: str) -> Tuple[Dict[str, SectionSlice], str]:
    raw_sections, mode = split_sections(full_text)
    out: Dict[str, SectionSlice] = {}
    for key, raw in raw_sections.items():
        own, ratio = extract_own_content(raw)
        out[key] = SectionSlice(
            key=key,
            raw=raw,
            own=own,
            word_count_own=_word_count(own),
            consigna_ratio=ratio,
        )
    return out, mode


def _global_own_text(full_text: str) -> str:
    own, _ = extract_own_content(full_text)
    return own


def _combined_own_text(
    slices: Dict[str, SectionSlice],
    keys: List[str],
    full_text: str,
    min_words_fallback: int = 50,
) -> Tuple[str, str]:
    parts = []
    used_keys = []
    for k in keys:
        if k == "cuerpo":
            continue
        sl = slices.get(k)
        if sl and sl.own.strip():
            parts.append(sl.own)
            used_keys.append(k)
    combined = "\n\n".join(parts)
    if _word_count(combined) >= min_words_fallback:
        return combined, "apartados"
    global_own = _global_own_text(full_text)
    if _word_count(global_own) > _word_count(combined):
        return global_own, "texto_global"
    return combined, "apartados" if combined.strip() else "texto_global"


def _check_presence(own_text: str, terms: List[str]) -> Tuple[int, int, List[str]]:
    low = own_text.lower()
    found = []
    for t in terms:
        if t.lower() in low:
            found.append(t)
    return len(found), len(terms), found


def _length_fraction(word_own: int, keys: List[str], slices: Dict[str, SectionSlice]) -> float:
    mins, maxs = [], []
    for sk in keys:
        if sk == "cuerpo":
            continue
        if sk in SECTION_WORD_RANGES:
            a, b = SECTION_WORD_RANGES[sk]
            mins.append(a)
            maxs.append(b)
    if not mins:
        w_min, w_max = 120, 600
    else:
        w_min = int(sum(mins) / len(mins))
        w_max = int(sum(maxs) / len(maxs))
    if word_own >= w_min:
        return min(1.0, 0.5 + 0.5 * min(1.0, word_own / max(w_max, w_min)))
    if word_own >= w_min * 0.35:
        return 0.35 + 0.45 * (word_own / max(1, w_min))
    return 0.2 * (word_own / max(1, w_min * 0.35)) if word_own else 0.0


def _bibliography_score(own_text: str) -> Tuple[float, List[str], List[str]]:
    notas: List[str] = []
    refs = re.findall(
        r"(?im)(?:[A-ZÁÉÍÓÚ][^\n,]{2,40},?\s*\(?\d{4}\)?|\(\d{4}\)|,\s*\d{4}\.)",
        own_text,
    )
    years = re.findall(r"\b(20(?:1[6-9]|2[0-6]))\b", own_text)
    if len(refs) < 2 and len(set(years)) < 2:
        notas.append("Pocas referencias bibliográficas en contenido propio.")
        return 0.0, [], notas
    frac = 0.4
    if len(refs) >= 5:
        frac += 0.3
    elif len(refs) >= 2:
        frac += 0.15
    recent = sum(1 for y in years if int(y) >= 2021)
    if recent >= 3:
        frac += 0.25
    elif recent >= 1:
        frac += 0.12
    return min(1.0, frac), [f"{len(refs)} referencias", f"{len(set(years))} años"], notas


def _presupuesto_score(own_text: str) -> Tuple[float, List[str], List[str]]:
    notas: List[str] = []
    low = own_text.lower()
    has_money = bool(re.search(r"(\$|usd|ars|pesos|presupuesto|costo|gasto|financiamiento)", low))
    has_numbers = bool(re.search(r"\b\d{2,}(?:[.,]\d{3})*\b", own_text))
    has_sustain = any(x in low for x in ("sostenibilidad", "fuente", "financiamiento", "fondos"))
    if not ((has_money or has_numbers) and has_sustain):
        notas.append("Sin evidencia clara de presupuesto/financiamiento (Anexo IV → 0 en este ítem).")
        return 0.0, [], notas
    frac = 0.5 + (0.2 if has_numbers else 0) + (0.15 if has_money else 0) + (0.1 if has_sustain else 0)
    return min(1.0, frac), ["Presupuesto/financiamiento en contenido propio"], notas


def score_criterion_content(
    criterio: str,
    meta: dict,
    slices: Dict[str, SectionSlice],
    full_text: str,
    doc_mode: str,
) -> CriterionScore:
    peso = int(meta.get("peso", 0))
    section_keys = CRITERION_SECTIONS.get(criterio, [])
    checks = CRITERION_CHECKS.get(criterio, meta.get("pistas", []))

    own_text, fuente = _combined_own_text(slices, section_keys, full_text)
    word_own = _word_count(own_text)

    mins, maxs = [], []
    for sk in section_keys:
        if sk in SECTION_WORD_RANGES and sk != "cuerpo":
            mins.append(SECTION_WORD_RANGES[sk][0])
            maxs.append(SECTION_WORD_RANGES[sk][1])
    w_min = int(sum(mins) / len(mins)) if mins else 120
    w_max = int(sum(maxs) / len(maxs)) if maxs else 600

    notas: List[str] = [f"Modo documento: {doc_mode}. Fuente de texto: {fuente}."]

    if not own_text.strip():
        notas.append("Sin contenido propio detectado para este criterio.")
        return CriterionScore(0, peso, 0, len(checks), 0, w_min, w_max, [], notas, doc_mode)

    if criterio == "Presupuesto y sostenibilidad":
        frac, ev, n = _presupuesto_score(own_text)
        notas.extend(n)
        return CriterionScore(
            round(peso * frac), peso, int(frac * 4), 4, word_own, w_min, w_max, ev, notas, doc_mode
        )

    if criterio == "Bibliografía actualizada":
        frac, ev, n = _bibliography_score(own_text)
        notas.extend(n)
        return CriterionScore(
            round(peso * frac), peso, len(ev), max(4, len(checks)), word_own, w_min, w_max, ev, notas, doc_mode
        )

    checks_ok, checks_total, found = _check_presence(own_text, list(checks))
    item_frac = checks_ok / max(1, checks_total)
    length_frac = _length_fraction(word_own, section_keys, slices)

    combined = (item_frac * 0.55) + (length_frac * 0.45)
    combined = min(1.0, combined)

    puntaje = min(peso, max(0, round(peso * combined)))

    evidencias = []
    for term in found[:3]:
        idx = own_text.lower().find(term.lower())
        if idx >= 0:
            evidencias.append(own_text[max(0, idx - 60) : idx + 140].replace("\n", " ")[:220])
    if not evidencias and own_text:
        evidencias.append(own_text[:220].replace("\n", " "))

    notas.append(f"Ítems sustantivos: {checks_ok}/{checks_total}. Palabras propias: {word_own}.")

    return CriterionScore(
        puntaje=puntaje,
        peso_max=peso,
        checks_ok=checks_ok,
        checks_total=checks_total,
        word_own=word_own,
        word_target_min=w_min,
        word_target_max=w_max,
        evidencias=evidencias[:2],
        notas=notas,
        modo=doc_mode,
    )


def score_project(
    full_text: str, criterios: Dict[str, dict]
) -> Tuple[Dict[str, CriterionScore], Dict[str, SectionSlice], str]:
    slices, mode = build_section_slices(full_text)
    results: Dict[str, CriterionScore] = {}
    for nombre, meta in criterios.items():
        results[nombre] = score_criterion_content(nombre, meta, slices, full_text, mode)
    return results, slices, mode
