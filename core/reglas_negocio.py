"""
Reglas de negocio del Centro de Idiomas: horas, tarifas y umbrales que antes
vivían como números sueltos —sin nombre ni comentario— repartidos por el
código (`core/planillas/functions.py`, `core/control_pagos/control_pagos.py`,
`core/fases/services/excel_reader_service.py`, `core/fases/services/
description_service.py`). Si el centro cambia una tarifa o un umbral, este
es el único lugar que hay que tocar — y el único lugar donde hace falta
buscar para saber de dónde sale cada número.

No reemplaza `core/correos/config.py`, que ya cumplía este mismo rol para
todo lo relacionado a correos (EmailConfig, GmailConfig, etc.); este módulo
cubre las reglas de planillas, pagos y generación de documentos.
"""

from dataclasses import dataclass
from typing import Optional

# ===== Horas por tipo de servicio =====

# Cada curso dictado (regular o superintensivo) equivale a esta cantidad de
# horas de dictado. Se usa tanto para calcular el monto a pagar como para
# describir el servicio en los documentos generados (oficios, TDR,
# cotizaciones, conformidades). Junior y Suficiencia tienen sus propias
# reglas: ver `regla_para_curso` más abajo.
HORAS_POR_CURSO_DICTADO = 28

# Cada curso dictado también genera esta cantidad de horas de diseño de
# exámenes, pagadas aparte.
HORAS_DISENO_EXAMENES_POR_CURSO = 4

# ===== Reglas por programa =====

@dataclass(frozen=True)
class ReglaPrograma:
    """Horas y tarifa que se pagan por CADA curso de un programa."""
    nombre: str
    horas_dictado: int
    horas_diseno_examenes: int
    # None = se paga según la categoría del docente. Un valor = tarifa por
    # hora fija para todos los docentes, sin importar su categoría.
    tarifa_hora_fija: Optional[float] = None


PROGRAMA_REGULAR = ReglaPrograma("Regular", HORAS_POR_CURSO_DICTADO, HORAS_DISENO_EXAMENES_POR_CURSO)

# Programa Junior: la mitad de horas de dictado, mismo diseño de exámenes.
PROGRAMA_JUNIOR = ReglaPrograma("Junior", 14, 4)

# Preparación para el examen de suficiencia: sin diseño de exámenes y a
# S/ 60 la hora para todos -> 28 h x S/ 60 = S/ 1680 por curso.
PROGRAMA_SUFICIENCIA = ReglaPrograma("Suficiencia", 28, 0, tarifa_hora_fija=60.0)


def regla_para_curso(texto_curso: str) -> ReglaPrograma:
    """
    Decide la regla de pago a partir del texto del curso (y del programa, si
    se concatena). En el sistema el idioma ya trae el programa: "Inglés
    Junior Básico 1", "Inglés Preparación Suficiencia A2 General 1".
    """
    texto = str(texto_curso or "").upper()
    if "SUFICIENCIA" in texto:
        return PROGRAMA_SUFICIENCIA
    if "JUNIOR" in texto:
        return PROGRAMA_JUNIOR
    return PROGRAMA_REGULAR

# ===== Tarifas =====

# Tarifa fija en soles por hora de servicio de coordinación/apoyo docente,
# usada SOLO cuando el docente no tiene una categoría de pago asignada en la
# base de docentes (caso excepcional; lo normal es pagar según su categoría).
TARIFA_COORDINACION_SIN_CATEGORIA = 50.0

# ===== Coincidencia de nombres (fuzzy matching) =====

# Puntaje mínimo (0-100, según fuzzywuzzy) para aceptar que un nombre de
# docente del archivo de cursos corresponde a un docente de la base de
# docentes al armar la planilla. Por debajo de este umbral se considera
# "sin coincidencia": se avisa en pantalla en vez de asignarlo en silencio.
UMBRAL_COINCIDENCIA_DOCENTE_PLANILLA = 65

# Igual que el anterior, pero al emparejar los nombres del archivo de
# control de pagos contra los de la planilla. Es más alto porque acá un
# falso positivo mueve dinero de armada a la persona equivocada.
UMBRAL_COINCIDENCIA_DOCENTE_CONTROL_PAGOS = 85

# ===== Control de pagos =====

# Saldo restante (en soles) por debajo del cual se avisa que el contrato de
# un docente está por terminar.
SALDO_MAXIMO_CONTRATO_POR_ACABAR = 1280.0
