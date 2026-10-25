"""Domain models for docentes, pagos, cursos, and document configuration."""

from .docente_data import DocenteData
from .payment_data import PaymentData
from .document_config import DocumentConfig, TipoDocumento, TipoContrato
from .curso_detalle import CursoDetalle
from .resultado_generacion import ResultadoDocente, ResultadoProcesoFase

__all__ = [
    'DocenteData',
    'PaymentData',
    'DocumentConfig',
    'TipoDocumento',
    'TipoContrato',
    'CursoDetalle',
    'ResultadoDocente',
    'ResultadoProcesoFase',
]
