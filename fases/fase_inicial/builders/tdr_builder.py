import os
from typing import Dict
from docx2pdf import convert
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import PaymentService, DescriptionService, DocumentGeneratorService
from fases.utils import PathUtils


class TdrBuilder:
    
    def __init__(self, config: DocumentConfig):
        self.config = config
        self.payment_service = PaymentService()
        self.description_service = DescriptionService()
        self.doc_service = DocumentGeneratorService()
    
    def construir_reemplazos(
        self,
        docente: DocenteData,
        payment: PaymentData,
        descripcion_completa: str
    ) -> Dict[str, str]:
        montos = self.payment_service.calcular_montos_completos(
            payment, 
            self.config.es_administrativo
        )
        
        # Generar actividades según el tipo
        actividades = self.description_service.generar_actividades_docentes(payment)
        
        # Generar monto referencial
        monto_referencial = self.payment_service.generar_monto_referencial(
            montos['monto_sin_actualizacion'],
            montos['monto_sin_actualizacion_letras'],
            montos['servicio_actualizacion'],
            montos['servicio_actualizacion_letras'],
            montos['bono_para_mostrar'],
            montos['bono_letras'],
            montos['monto_total'],
            montos['monto_total_letras']
        )
        
        reemplazos = {
            "descripcion": descripcion_completa,
            "actividades_docentes": actividades,
            "categoria": montos['categoria_formato'],
            "monto_subtotal": monto_referencial,
            "modalidad_servicio": self.config.modalidad_servicio,
        }
        
        # Agregar campos adicionales para administrativos
        if self.config.es_administrativo:
            reemplazos.update({
                "finalidad_publica": docente.finalidad_publica,
                "formacion_academica": docente.formacion_academica,
                "experiencia_laboral": docente.experiencia_laboral,
                "requisitos_adicional": docente.requisitos_adicional,
                "actividades_admin": docente.curso  # Para administrativos, curso contiene actividades
            })
        
        return reemplazos
    
    def generar(
        self,
        docente: DocenteData,
        payment: PaymentData,
        descripcion_completa: str,
        carpeta_docente: str
    ) -> str:
        # Obtener plantilla
        if self.config.es_administrativo:
            ruta_template = PathUtils.ruta_absoluta_relativa('./Modelos_documentos/tdr_administrativo.docx')
        else:
            # Para docentes, usar plantilla según categoría
            ruta_template = PathUtils.ruta_absoluta_relativa(
                f'./Modelos_documentos/tdr_tipo{docente.categoria_letra}_.docx'
            )
        
        # Construir reemplazos
        reemplazos = self.construir_reemplazos(docente, payment, descripcion_completa)
        
        # Generar nombre de archivo
        nombre_archivo = f"TDR - {docente.nombre_limpio} - {self.config.mes} {self.config.anio}.docx"
        ruta_salida = os.path.join(carpeta_docente, nombre_archivo)
        
        # Generar documento
        self.doc_service.generar_documento(ruta_template, reemplazos, ruta_salida)
        
        # Limpiar texto administrativo si aplica
        if self.config.es_administrativo:
            self.doc_service.eliminar_texto_administrativo(ruta_salida)
        
        # Convertir a PDF
        try:
            ruta_pdf = PathUtils.cambiar_extension(ruta_salida, 'pdf')
            convert(ruta_salida, ruta_pdf)
        except Exception as e:
            print(f"Error al convertir TDR a PDF: {e}")
        
        return ruta_salida
