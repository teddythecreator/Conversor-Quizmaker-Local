# Conversor DOCX a XLSX para Quiz Maker (WordPress Plugin) - Versión Streamlit
# Autor: Tedi One - Nexo de Negocios Digitales

import docx
import pandas as pd
import streamlit as st
from io import BytesIO
import json

EXPLICACION_TEXTO = "Por favor revisa la explicación de la respuesta para entender mejor el tema abordado."
TIPO_PREGUNTA = "radio"

# === FUNCIONES ===
def cargar_documento(file):
    doc = docx.Document(file)
    return [p for p in doc.paragraphs if p.text.strip() != ""]

def es_respuesta_correcta(run):
    try:
        return run.bold or (hasattr(run.font, 'highlight_color') and run.font.highlight_color is not None)
    except:
        return False
    except:
        return False

def extraer_preguntas_y_respuestas(parrafos):
    preguntas = []
    i = 0
    while i < len(parrafos):
        texto = parrafos[i].text.strip() if hasattr(parrafos[i], 'text') else parrafos[i]

        if len(texto.split()) > 3 and not texto.lower().startswith(("respuesta", "examen", "plantilla", "explicación")):
            pregunta = texto
            respuestas = []
            explicacion = ""
            i += 1
            while i < len(parrafos):
                p = parrafos[i]
                p_text = p.text.strip() if hasattr(p, 'text') else p.strip()

                # Detectar explicación por texto completamente en negrita o resaltado
                if hasattr(p, 'runs') and any(es_respuesta_correcta(run) for run in p.runs):
                    if all(es_respuesta_correcta(run) for run in p.runs if run.text.strip()) and len(p_text.split()) <= 3:
                        explicacion = p_text
                        i += 1
                        continue

                elif len(p_text) == 0:
                    i += 1
                    continue
                else:
                    respuesta = ""
                    correcta = False
                    if hasattr(p, 'runs'):
                        for run in p.runs:
                            if es_respuesta_correcta(run):
                                correcta = True
                            respuesta += run.text
                    else:
                        if p_text.startswith("*") or p_text.startswith("✔") or "(*)" in p_text:
                            correcta = True
                            respuesta = p_text.lstrip("*✔ ").replace("(*)", "").strip()
                        else:
                            respuesta = p_text
                    respuestas.append((respuesta.strip(), correcta))
                i += 1

            if not explicacion:
                explicacion = EXPLICACION_TEXTO
            if len(respuestas) < 2:
                st.warning(f"❗ La pregunta '{pregunta}' tiene menos de 2 respuestas. Será omitida.")
            elif not any(c for _, c in respuestas):
                st.warning(f"⚠️ La pregunta '{pregunta}' no tiene ninguna respuesta marcada como correcta.")
            else:
                preguntas.append({
                    "pregunta": pregunta,
                    "respuestas": respuestas,
                    "explicacion": explicacion
                })
        else:
            i += 1
    return preguntas

def construir_estructura_xlsx(preguntas):
    data = []
    for item in preguntas:
        answers_json = []
        for i, (texto, correcto) in enumerate(item["respuestas"], start=1):
            answers_json.append({
                "id": "",
                "question_id": "",
                "answer": texto,
                "image": "",
                "correct": "1" if correcto else "0",
                "ordering": str(i),
                "weight": "1",
                "keyword": "",
                "placeholder": "",
                "slug": "",
                "options": ""
            })

        fila = {
            "id": "",
            "category": "",
            "question": item["pregunta"],
            "question_title": "",
            "question_image": "",
            "question_hint": "",
            "type": TIPO_PREGUNTA,
            "published": "1",
            "wrong_answer_text": item["explicacion"],
            "right_answer_text": item["explicacion"],
            "explanation": "",
            "user_explanation": "off",
            "not_influence_to_score": "off",
            "weight": "1",
            "options": "",
            "question_id": "",
            "tags": "",
            "answers": json.dumps(answers_json, ensure_ascii=False)
        }
        data.append(fila)
    return pd.DataFrame(data)

def convertir_y_descargar(uploaded_file):
    parrafos = cargar_documento(uploaded_file)
    preguntas = extraer_preguntas_y_respuestas(parrafos)
    if not preguntas:
        raise ValueError("No se encontraron preguntas válidas en el archivo.")
    df = construir_estructura_xlsx(preguntas)
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

# === INTERFAZ STREAMLIT ===
st.title("Conversor DOCX a XLSX - Quiz Maker (Formato Avanzado)")
st.markdown("Sube tu archivo .docx con preguntas tipo test y descarga un archivo .xlsx listo para importar en el plugin WordPress Quiz Maker (formato completo).")

uploaded_file = st.file_uploader("Selecciona el archivo DOCX", type=["docx"])

if uploaded_file:
    if st.button("Convertir y descargar XLSX"):
        try:
            xlsx_data = convertir_y_descargar(uploaded_file)
            st.success("Conversión completada. Descarga el archivo a continuación.")
            st.download_button(
                label="📥 Descargar archivo XLSX",
                data=xlsx_data,
                file_name="preguntas_quiz.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
