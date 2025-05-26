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
        # Detectar pregunta aunque no termine en ?
        if len(texto.split()) > 3:
            pregunta = texto
            respuestas = []
            explicacion = ""
            respuesta_correcta_letra = ""
            i += 1
            while i < len(parrafos):
                line = parrafos[i].text.strip()
                line_lower = line.lower()
                # Detección flexible de la línea de respuesta correcta
                if "respuesta correcta" in line_lower or "respuesta:" in line_lower:
                    match = re.search(r"[a-d]", line_lower)
                    if match:
                        respuesta_correcta_letra = match.group(0).lower()
                    i += 1
                elif "explicación correcta" in line_lower or "explicacion correcta" in line_lower:
                    explicacion = re.sub(r"explicaci[oó]n correcta[:]*", "", line, flags=re.IGNORECASE).strip()
                    i += 1
                    break
                elif line == "":
                    i += 1
                    continue
                else:
                    # Limpiar espacios y numeración tipo "1.", "-", etc.
                    respuesta_limpia = re.sub(r"^[0-9]+\.|-", "", line).strip()
                    respuestas.append(respuesta_limpia)
                    i += 1
            # Vincular la letra de respuesta correcta con el texto
            respuestas_finales = []
            for idx, texto_r in enumerate(respuestas):
                letra_opcion = letras[idx] if idx < len(letras) else ""
                es_correcta = (letra_opcion == respuesta_correcta_letra)
                respuestas_finales.append((texto_r, es_correcta))
            if len(respuestas_finales) >= 2:
                preguntas.append({
                    "pregunta": pregunta,
                    "respuestas": respuestas_finales,
                    "explicacion": explicacion or EXPLICACION_TEXTO
                })
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
st.title("Conversor DOCX a XLSX - Quiz Maker (Formato Flexible)")
st.markdown("Sube tu archivo .docx con preguntas tipo test (formato flexible) y descarga un archivo .xlsx listo para importar en WordPress Quiz Maker.")

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
