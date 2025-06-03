import re
import pandas as pd
from num2words import num2words

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", nombre)


def limpiar_numero(valor):
    return "" if pd.isna(valor) else str(valor).split('.')[0]


def reemplazar_en_parrafos(documento, reemplazos):
    for parrafo in documento.paragraphs:
        for marcador, valor in reemplazos.items():
            if marcador in parrafo.text:
                texto_nuevo = parrafo.text.replace(marcador, valor)
                for run in parrafo.runs:
                    run.text = ''
                if parrafo.runs:
                    parrafo.runs[0].text = texto_nuevo

def reemplazar_en_tablas(documento, reemplazos):
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for marcador, valor in reemplazos.items():
                        if marcador in parrafo.text:
                            texto_nuevo = parrafo.text.replace(marcador, valor)
                            for run in parrafo.runs:
                                run.text = ''
                            if parrafo.runs:
                                parrafo.runs[0].text = texto_nuevo

def monto_a_letras(monto):
    try:
        monto = float(monto)
        entero = int(monto)
        centavos = int(round((monto - entero) * 100))
        return f"{num2words(entero, lang='es')} y {centavos:02d}/100 soles"
    except Exception:
        return "N/A"

def redactar_cursos(cadena):
    if not isinstance(cadena, str):
        return "N/A"
    cursos = [c.strip() for c in cadena.split("/") if c.strip()]
    if not cursos:
        return "N/A"
    resultado = f"servicio de dictado de 28 horas de clases de {cursos[0]}"
    for curso in cursos[1:]:
        resultado += f", 28 horas de clases de {curso}"
    return resultado