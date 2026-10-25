"""Helpers de formateo de datos sueltos (días, horas, números) sin lógica de negocio."""

import re
import pandas as pd


def traducir_dias(dias_raw: str) -> str:
    dias_dict = {'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié', 'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'}
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)


def extraer_numero_horas(texto):
    if pd.isna(texto):
        return 0
    numeros = re.findall(r'\d+', str(texto))
    return int(numeros[0]) if numeros else 0


def formatear_numero(valor, ancho_minimo):
    if pd.isna(valor):
        return ''

    valor_str = str(valor).strip()
    if not valor_str or valor_str.lower() == 'nan':
        return ''

    if '.' in valor_str:
        try:
            valor_str = str(int(float(valor_str)))
        except (ValueError, TypeError):
            pass

    if valor_str.isdigit():
        return valor_str.zfill(max(ancho_minimo, len(valor_str)))

    return valor_str
