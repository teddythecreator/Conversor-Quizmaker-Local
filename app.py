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
            respuesta_correcta_texto = ""
            i += 1
            # Recogemos las 4 respuestas (ignorando líneas vacías)
            while i < len(parrafos) and len(respuestas) < 4:
                line = parrafos[i].text.strip()
                if line and not line.lower().startswith(("respuesta correcta", "explicación correcta", "explicacion correcta")):
                    respuestas.append(line)
                i += 1
            # Buscamos la respuesta correcta y la explicación
            while i < len(parrafos):
                line_lower = parrafos[i].text.strip().lower()
                if "respuesta correcta" in line_lower:
                    match = re.search(r"[a-d]", line_lower)
                    if match:
                        idx = ord(match.group(0)) - ord('a')
                        if 0 <= idx < len(respuestas):
                            respuesta_correcta_texto = respuestas[idx]
                    i += 1
                elif "explicación correcta" in line_lower or "explicacion correcta" in line_lower:
                    explicacion = re.sub(r"explicaci[oó]n correcta[:]*", "", parrafos[i].text.strip(), flags=re.IGNORECASE).strip()
                    i += 1
                    break
                else:
                    i += 1
            # Vinculamos la respuesta correcta comparando minúsculas y sin espacios
            respuestas_finales = [(texto_r, texto_r.strip().lower() == respuesta_correcta_texto.strip().lower()) for texto_r in respuestas]
            if len(respuestas_finales) >= 2:
                preguntas.append({
                    "pregunta": pregunta,
                    "respuestas": respuestas_finales,
                    "explicacion": explicacion or EXPLICACION_TEXTO
                })
        else:
            i += 1
    return preguntas
