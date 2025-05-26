# Conversor DOCX a XLSX para Quiz Maker (WordPress Plugin) - Versión Streamlit
# Autor: Tedi One - Nexo de Negocios Digitales

import docx
import pandas as pd
import streamlit as st
from io import BytesIO
import json
import re

EXPLICACION_TEXTO = "Por favor revisa la explicación de la respuesta para entender mejor el tema abordado."
TIPO_PREGUNTA = "radio"

def cargar_documento(file):
    doc = docx.Document(file)
    return [p for p in doc.paragraphs if p.text.strip() != ""]

def extraer_preguntas_y_respuestas(parrafos):
    preguntas = []
    i = 0
    letras = ["a", "b", "c", "d"]
    while i < len(parrafos):
        texto = parrafos[i].text.strip()
        if len(texto.split()) > 3:
            pregunta = texto
            respuestas = []
            explicacion = ""
            respuesta_correcta = ""
            i += 1
            # Recogemos las 4 respuestas REALES (ignorando líneas vacías o solo espacios)
            while i < len(parrafos) and len(respuestas) < 4:
                line = parrafos[i].text.strip()
                if line and not line.lower().startswith(("respuesta correcta", "explicación correcta", "explicacion correcta")):
                    respuestas.append(line)
                i += 1
            # Buscamos la línea de respuesta correcta y la explicación correcta
            while i < len(parrafos):
                line = parrafos[i].text.strip().lower()
                if "respuesta correcta" in line:
                    match = re.search(r"[a-d]", line)
                    if match:
                        letra_idx = match.group(0).lower()
                        idx = ord(letra_idx) - ord('a')
                        if 0 <= idx < len(respuestas):
                            respuesta_correcta = respuestas[idx]
                    i += 1
                elif "explicación correcta" in line or "explicacion correcta" in line:
                    explicacion = re.sub(r"explicaci[oó]n correcta[:]*", "", parrafos[i].text.strip(), flags=re.IGNORECASE).strip()
                    i += 1
                    break
                else:
                    i += 1
            # Vinculamos la respuesta correcta con el texto real
            respuestas_finales = [(texto_r, texto_r == respuesta_correcta) for texto_r in respuestas]
            if len(respuestas_finales) >= 2:
                preguntas.append({
                    "pregunta": pregunta,
                    "respuestas": respuestas_finales,
                    "explicacion": explicacion or EXPLICACION_TEXTO
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
st.title("Conversor DOCX a XLSX - Quiz Maker (Final y 100%)")
st.markdown("Sube tu archivo .docx con preguntas y descarga el archivo .xlsx listo para importar en el plugin WordPress Quiz Maker.")

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
