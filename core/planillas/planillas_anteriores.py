"""
Comparación contra las planillas ya emitidas en el mes.

Normalmente se emiten dos planillas por mes: la primera con la mayoría de
docentes y una segunda, más chica, con lo que faltó. Un curso se identifica
por la pareja (ID Programación, docente):

- mismo ID y mismo docente que en una planilla anterior -> ya se emitió, se excluye;
- mismo ID pero otro docente -> cambió de docente, se incluye;
- ID que no estaba -> curso nuevo, se incluye.

Antes esta comparación se hacía a mano entre los dos CSV.
"""

import os
from dataclasses import dataclass, field
from typing import Iterable, List, Set, Tuple

import pandas as pd

from core.fases.utils import TextUtils
from core.planillas import columnas as col
from core.planillas.csv_processing import (
    cargar_archivo,
    filtrar_docentes_validos,
    limpiar_docentes,
    normalizar_columnas_planilla,
)

RESULTADO_YA_EMITIDO = "Excluido: ya emitido"
RESULTADO_CAMBIO_DOCENTE = "Incluido: cambio de docente"
RESULTADO_NUEVO = "Incluido: curso nuevo"


@dataclass
class ResumenComparacion:
    planillas_anteriores: int
    excluidos: int = 0
    cambio_docente: int = 0
    nuevos: int = 0
    # Una fila por curso del CSV actual con el resultado de la comparación,
    # para la hoja "Comparación" del Excel.
    detalle: pd.DataFrame = field(default_factory=pd.DataFrame)


def _normalizar_id(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    # pandas puede leer el ID como número (104386.0).
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]
    return texto


def _validar_columna_id(datos: pd.DataFrame, ruta: str) -> None:
    if col.ID_PROGRAMACION not in datos.columns:
        raise ValueError(
            f"El archivo {os.path.basename(ruta)} no tiene la columna 'ID Programación', "
            "necesaria para compararlo con las planillas anteriores del mes."
        )


def _parejas(datos: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "id": datos[col.ID_PROGRAMACION].apply(_normalizar_id),
        "docente": datos[col.DOCENTE_ORIGINAL].apply(TextUtils.normalizar_texto),
    }, index=datos.index)


def cargar_cursos_emitidos(rutas_csv: Iterable[str]) -> Tuple[Set[Tuple[str, str]], dict]:
    """
    Lee los CSV de las planillas anteriores y devuelve:
    - el conjunto de parejas (ID, docente normalizado) ya emitidas;
    - por cada ID, los nombres de docente con los que se emitió (para el detalle).
    """
    emitidos: Set[Tuple[str, str]] = set()
    docentes_por_id: dict = {}

    for ruta in rutas_csv:
        datos = normalizar_columnas_planilla(cargar_archivo(ruta))
        _validar_columna_id(datos, ruta)
        datos = limpiar_docentes(datos, col.DOCENTE_ORIGINAL)
        datos = filtrar_docentes_validos(datos, col.DOCENTE_ORIGINAL)

        parejas = _parejas(datos)
        for (id_prog, docente_norm), docente in zip(parejas.itertuples(index=False), datos[col.DOCENTE_ORIGINAL]):
            if not id_prog:
                continue
            emitidos.add((id_prog, docente_norm))
            docentes_por_id.setdefault(id_prog, [])
            if docente not in docentes_por_id[id_prog]:
                docentes_por_id[id_prog].append(docente)

    return emitidos, docentes_por_id


def excluir_cursos_emitidos(datos: pd.DataFrame, rutas_anteriores: List[str], ruta_actual: str) -> Tuple[pd.DataFrame, ResumenComparacion]:
    """Quita de `datos` (CSV de la planilla actual) los cursos que ya salieron en `rutas_anteriores`."""
    _validar_columna_id(datos, ruta_actual)
    emitidos, docentes_por_id = cargar_cursos_emitidos(rutas_anteriores)

    parejas = _parejas(datos)
    ya_emitido = pd.Series(
        [(id_prog, docente) in emitidos for id_prog, docente in parejas.itertuples(index=False)],
        index=datos.index,
    )
    id_existia = parejas["id"].isin(docentes_por_id.keys()) & (parejas["id"] != "")

    resultado = pd.Series(RESULTADO_NUEVO, index=datos.index)
    resultado[id_existia & ~ya_emitido] = RESULTADO_CAMBIO_DOCENTE
    resultado[ya_emitido] = RESULTADO_YA_EMITIDO

    detalle = pd.DataFrame({
        "ID Programación": parejas["id"],
        "Curso": datos["Curso"] if "Curso" in datos.columns else "",
        "Docente": datos[col.DOCENTE_ORIGINAL],
        "Docente en planilla anterior": parejas["id"].map(lambda i: " / ".join(docentes_por_id.get(i, []))),
        "Resultado": resultado,
    })
    detalle = detalle.sort_values(["Resultado", "Docente"]).reset_index(drop=True)
    detalle.insert(0, "N°", range(1, len(detalle) + 1))

    resumen = ResumenComparacion(
        planillas_anteriores=len(rutas_anteriores),
        excluidos=int(ya_emitido.sum()),
        cambio_docente=int((resultado == RESULTADO_CAMBIO_DOCENTE).sum()),
        nuevos=int((resultado == RESULTADO_NUEVO).sum()),
        detalle=detalle,
    )
    return datos[~ya_emitido].copy(), resumen
