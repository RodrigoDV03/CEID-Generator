"""
Resultado explícito de generar documentos para uno o varios docentes.

Reemplaza el patrón de "imprimir ✅/❌ y devolver None", que obligaba a la
interfaz a asumir que todo salió bien salvo que una excepción se escapara.
Con esto, la pantalla puede mostrar un resumen real (cuántos se generaron,
quiénes quedaron sin firma, quiénes fallaron) en vez de un genérico
"Documentos generados correctamente" aunque algo haya salido mal.
"""

from dataclasses import dataclass, field
from typing import List


@dataclass
class ResultadoDocente:
    """Resultado de generar los documentos de un docente/administrativo puntual."""
    nombre: str
    ok: bool
    firma_faltante: bool = False
    error: str = ""


@dataclass
class ResultadoProcesoFase:
    """Resumen de correr `procesar_planilla_fase_inicial` para toda la planilla."""
    total_procesados: int = 0
    exitosos: int = 0
    con_error: List[str] = field(default_factory=list)
    sin_firma: List[str] = field(default_factory=list)

    def registrar(self, resultado: ResultadoDocente) -> None:
        self.total_procesados += 1
        if resultado.ok:
            self.exitosos += 1
        else:
            self.con_error.append(resultado.nombre)
        if resultado.firma_faltante:
            self.sin_firma.append(resultado.nombre)

    @property
    def tiene_advertencias(self) -> bool:
        return bool(self.con_error or self.sin_firma)
