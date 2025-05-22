# 🧠 Conversor DOCX a XLSX para WordPress Quiz Maker

**Autor**: Tedi One – Nexo de Negocios Digitales  
**Versión**: Streamlit App (Avanzado)

---

## 📌 Descripción

Esta herramienta convierte automáticamente archivos `.docx` que contienen preguntas tipo test al formato `.xlsx` compatible con el plugin **Quiz Maker para WordPress**.

Está diseñada para importar preguntas de manera masiva sin necesidad de ingresarlas una a una manualmente.

---

## ⚙️ Características

- ✅ Compatible con formato avanzado de importación de Quiz Maker.
- ✅ Detecta preguntas y respuestas desde `.docx`.
- ✅ Identifica la respuesta correcta por **negrita o resaltado**.
- ✅ Extrae explicación automática desde líneas completamente en negrita.
- ✅ Genera archivo `.xlsx` listo para importar.
- ❌ PDF no soportado en esta versión (por estabilidad).

---

## 📂 Estructura de los datos generados

Cada pregunta exportada incluye:

- `question` — Texto de la pregunta
- `answers` — Respuestas en formato JSON
- `correct` — Identificador binario (`1` o `0`)
- `wrong_answer_text` y `right_answer_text` — Texto de explicación
- `type` — Siempre `"radio"`
- Otras columnas requeridas por el plugin: `published`, `weight`, etc.

---

## 🚀 Cómo usar

1. Ejecuta la aplicación con Streamlit:

   ```bash
   streamlit run app.py
