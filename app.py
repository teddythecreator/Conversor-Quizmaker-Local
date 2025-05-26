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
    return [p.text.strip() for p in doc.paragraphs if p.text.strip() != ""]

def detectar_preguntas(parrafos):
    bloques = []
    i = 0
    while i < len(parrafos):
        texto = parrafos[i]
        if texto and not texto.lower().startswith(("respuesta correcta", "explicación correcta", "explicacion correcta")):
            pregunta = texto
            respuestas = []
            i += 1
            while i < len(parrafos) and len(respuestas) < 4:
                line = parrafos[i].strip()
                if line and not line.lower().startswith(("respuesta correcta", "explicación correcta", "explicacion correcta")):
                    respuestas.append(line)
                i += 1
            bloques.append({
                "pregunta": pregunta,
                "respuestas": respuestas,
                "indice": i  # Para continuar buscando desde aquí
            })
        else:
            i += 1
    return bloques

def detectar_respuesta_correcta(parrafos, start_idx):
    respuesta_letra = ""
    i = start_idx
    while i < len(parrafos):
        line = parrafos[i].lower()
        if "respuesta correcta" in line:
            match = re.search(r"[a-d]", line)
            if match:
                respuesta_letra = match.group(0).lower()
            break
        i += 1
    return respuesta_letra

def detectar_explicacion(parrafos, start_idx):
    i = start_idx
    while i < len(parrafos):
        line = parrafos[i].lower()
        if "explicación correcta" in line or "explicacion correcta" in line:
            explicacion = re.sub(r"explicaci[oó]n correcta[:]*", "", parrafos[i], flags=re.IGNORECASE).strip()
            return explicacion
        i += 1
    return EXPLICACION_TEXTO

def extraer_preguntas_completas(parrafos):
    bloques = detectar_preguntas(parrafos)
    preguntas_completas = []

    for bloque in bloques:
        pregunta = bloque["pregunta"]
        respuestas = bloque["respuestas"]
        start_idx = bloque["indice"]

        letra_correcta = detectar_respuesta_correcta(parrafos, start_idx)
        idx_correcta = ord(letra_correcta) - ord('a') if letra_correcta else -1

        explicacion = detectar_explicacion(parrafos, start_idx)

        # Vinculamos la respuesta correcta real según la letra
        respuestas_finales = []
        for idx, texto in enumerate(respuestas):
            es_correcta = idx == idx_correcta
            respuestas_finales.append((texto, es_correcta))

        if len(respuestas_finales) >= 2:
            preguntas_completas.append({
                "pregunta": pregunta,
                "respuestas": respuestas_finales,
                "explicacion": explicacion
            })

    return preguntas_completas

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
    preguntas = extraer_preguntas_completas(parrafos)
    if not preguntas:
        raise ValueError("No se encontraron preguntas válidas en el archivo.")
    df = construir_estructura_xlsx(preguntas)
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

# === INTERFAZ STREAMLIT ===
st.title("Conversor DOCX a XLSX - Quiz Maker (Versión Modular y 100%)")
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
