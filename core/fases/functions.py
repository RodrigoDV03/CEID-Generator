def _iterar_parrafos(documento):
    for parrafo in documento.paragraphs:
        yield parrafo

    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    yield parrafo


def reemplazar_en_parrafo(parrafo, reemplazos):
    for marcador, valor in reemplazos.items():
        if marcador in parrafo.text:
            texto_nuevo = parrafo.text.replace(marcador, valor)
            for run in parrafo.runs:
                run.text = ''
            if parrafo.runs:
                parrafo.runs[0].text = texto_nuevo


def reemplazar_en_documento(documento, reemplazos):
    for parrafo in _iterar_parrafos(documento):
        reemplazar_en_parrafo(parrafo, reemplazos)
