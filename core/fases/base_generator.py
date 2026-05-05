from core.fases.models import DocenteData, PaymentData, DocumentConfig
from core.fases.services import (
    ExcelReaderService,
    DocenteService,
    PaymentService,
    DescriptionService,
)


class BaseFaseGenerator:
    def __init__(self, config: DocumentConfig):
        self.config = config

        # Servicios compartidos por los generadores de fase.
        self.excel_service = ExcelReaderService()
        self.docente_service = DocenteService()
        self.payment_service = PaymentService()
        self.description_service = DescriptionService()

    def generar_descripcion_completa(
        self,
        docente: DocenteData,
        payment: PaymentData,
        planilla_path: str,
    ) -> str:
        # Para administrativos, retornar el curso directamente.
        if self.config.es_administrativo:
            return docente.curso

        # Leer cursos detallados desde Planilla_Generador expandida.
        cursos_detallados = self.excel_service.leer_cursos_detallados_por_docente(
            planilla_path,
            "Planilla_Generador",
            docente.nombre,
        )

        if not cursos_detallados:
            # Fallback: usar método antiguo si no hay cursos detallados.
            descripcion_base = self.description_service.redactar_cursos(
                docente.curso,
                tiene_bono=payment.tiene_bono,
                tiene_servicio_actualizacion=payment.tiene_servicio_actualizacion,
            )

            # Generar descripciones de horas.
            horas_disenio, horas_clasif = self.payment_service.generar_descripcion_horas(payment)

            return self.description_service.generar_descripcion_completa(
                descripcion_base,
                payment,
                horas_disenio,
                horas_clasif,
                es_administrativo=self.config.es_administrativo,
            )

        # La modalidad debe incluirse en cada servicio para ambos tipos de docentes
        # (contratados y terceros)
        incluir_modalidad_por_item = True
        return self.description_service.redactar_servicios_con_modalidad(
            cursos_detallados,
            incluir_modalidad_por_item=incluir_modalidad_por_item,
        )