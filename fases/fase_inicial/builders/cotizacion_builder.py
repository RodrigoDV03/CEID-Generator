import os
from typing import Dict
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import PaymentService, DescriptionService, DocenteService, DocumentGeneratorService
from fases.utils import PathUtils


class CotizacionBuilder:
    
    def __init__(self, config: DocumentConfig):
        self.config = config
        self.payment_service = PaymentService()
        self.description_service = DescriptionService()
        self.docente_service = DocenteService()
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
        if self.config.es_administrativo:
            # Para administrativos, las actividades están en actividades_admin
            actividades = docente.actividades_admin
        else:
            # Para docentes, generar actividades con formato de guiones
            actividades = self.description_service.generar_actividades_admin_cotizacion(payment)
        
        # Obtener datos de contacto formateados
        contacto = self.docente_service.obtener_datos_contacto_formateados(docente)
        
        # Generar monto referencial con nota de impuestos
        monto_subtotal = self.payment_service.generar_monto_referencial(
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
            "nombre_docente": docente.nombre,
            "idioma_docente": docente.idioma,
            "descripcion_servicio": descripcion_completa,
            "actividades_admin": actividades,
            "categoria_monto": montos['categoria_formato'],
            "monto_subtotal": monto_subtotal,
            "modalidad_servicio": docente.modalidad_texto,
        }
        
        # Agregar datos de contacto
        reemplazos.update(contacto)
        
        return reemplazos
    
    def generar(
        self,
        docente: DocenteData,
        payment: PaymentData,
        descripcion_completa: str,
        carpeta_docente: str
    ) -> str:
        # Obtener plantilla
        ruta_template = PathUtils.ruta_absoluta_relativa('./Modelos_documentos/modelo_cotizacion.docx')
        
        # Construir reemplazos
        reemplazos = self.construir_reemplazos(docente, payment, descripcion_completa)
        
        # Generar nombre de archivo
        nombre_archivo = f"COTIZACIÓN - {docente.nombre_limpio} - {self.config.mes} {self.config.anio}.docx"
        ruta_salida = os.path.join(carpeta_docente, nombre_archivo)
        
        # Obtener ruta de firma
        ruta_firma = self.docente_service.obtener_ruta_firma(docente, self.config)
        
        # Generar documento con firma
        self.doc_service.generar_documento(ruta_template, reemplazos, ruta_salida, ruta_firma)
        
        # Limpiar texto administrativo si aplica
        if self.config.es_administrativo:
            self.doc_service.eliminar_texto_administrativo(ruta_salida)
        
        return ruta_salida
