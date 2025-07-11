import re
import pandas as pd
from num2words import num2words



def limpiar_numero(valor):
    return "" if pd.isna(valor) else str(valor).split('.')[0]

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