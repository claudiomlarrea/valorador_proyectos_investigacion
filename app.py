import streamlit as st
import pandas as pd
import io
from datetime import datetime

# -------- CONFIG --------
st.set_page_config(page_title="Valorador de Proyectos", layout="wide")

st.title("🧮 Valorador de Proyectos de Investigación")
st.caption("Versión estable – evaluación manual")

# -------- CRITERIOS --------
CRITERIOS = {
    "Pertinencia y relevancia": 10,
    "Claridad del problema y objetivos": 10,
    "Originalidad / aporte": 8,
    "Solidez metodológica": 14,
    "Calidad de datos / muestra": 10,
    "Factibilidad y cronograma": 8,
    "Consideraciones éticas": 6,
    "Impacto esperado": 8,
    "Plan de difusión": 6,
    "Presupuesto": 6,
    "Alineación institucional": 6,
    "Bibliografía": 8,
}

# -------- EVALUACIÓN --------
st.subheader("Evaluación")

scores = {}
total_max = sum(CRITERIOS.values())

cols = st.columns(2)

i = 0
for criterio, peso in CRITERIOS.items():
    with cols[i % 2]:
        st.markdown(f"**{criterio}** (máx {peso})")
        val = st.slider(
            f"Puntaje {criterio}",
            0,
            peso,
            peso,  # arranca en máximo
            key=f"s_{i}"
        )
        scores[criterio] = val
        st.divider()
    i += 1

# -------- RESULTADO --------
total = sum(scores.values())
porcentaje = (total / total_max) * 100

def categoria(p):
    if p >= 70:
        return "Aprobado"
    elif p >= 50:
        return "Aprobado con observaciones"
    elif p >= 30:
        return "Requiere reformulación"
    else:
        return "No aprobado"

resultado = categoria(porcentaje)

st.markdown(f"## Resultado: **{resultado}**")
st.markdown(f"### Cumplimiento: **{round(porcentaje,2)}%**")

# -------- EXPORTAR EXCEL --------
def export_excel():
    df = pd.DataFrame(list(scores.items()), columns=["Criterio", "Puntaje"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

        resumen = pd.DataFrame([{
            "Resultado": resultado,
            "Porcentaje": round(porcentaje, 2),
            "Fecha": datetime.now()
        }])
        resumen.to_excel(writer, sheet_name="Resumen", index=False)

    return output.getvalue()

st.download_button(
    "⬇️ Descargar Excel",
    export_excel(),
    "resultado.xlsx"
)
