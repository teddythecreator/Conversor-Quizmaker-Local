def extraer_preguntas_y_respuestas(parrafos):
    preguntas = []
    i = 0
    while i < len(parrafos):
        texto = parrafos[i]
        # Detectar pregunta por numeración al inicio: "1.", "2.", etc.
        if re.match(r"^\\d+\\.", texto):
            pregunta = re.sub(r"^\\d+\\.", "", texto).strip()
            respuestas = []
            explicacion = ""
            respuesta_correcta_letra = ""
            i += 1
            while i < len(parrafos):
                line = parrafos[i].strip()
                line_lower = line.lower()
                if "respuesta correcta" in line_lower:
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
                    respuesta = re.sub(r"^[0-9]+\\.|-", "", line).strip()
                    respuestas.append(respuesta)
                    i += 1
            # Vincular la respuesta correcta real
            respuestas_finales = []
            letras = ["a", "b", "c", "d"]
            for idx, texto_r in enumerate(respuestas):
                letra = letras[idx] if idx < len(letras) else ""
                es_correcta = (letra == respuesta_correcta_letra)
                respuestas_finales.append((texto_r, es_correcta))
            if len(respuestas_finales) >= 2:
                preguntas.append({
                    "pregunta": pregunta,
                    "respuestas": respuestas_finales,
                    "explicacion": explicacion or EXPLICACION_TEXTO
                })
        i += 1
    return preguntas
