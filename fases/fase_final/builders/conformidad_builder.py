import os
from typing import Dict
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import PaymentService, DescriptionService, DocumentGeneratorService
from fases.utils import PathUtils


class ConformidadBuilder:
    
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
        
        # Asegurar que el número de contrato no esté vacío
        numero_contrato = docente.numero_contrato if docente.numero_contrato else "001"
        
        return {
            "nombre_docente": docente.nombre,
            "ruc": docente.ruc,
            "descripcion_cursos": descripcion_completa,
            "monto_subtotal": montos['monto_total_formato'],
            "monto_hora": montos['categoria_formato'],
            "Nro_Contrato": numero_contrato,
            "numero_armada": self.config.numero_armada,
            "modalidad_servicio": self.config.modalidad_servicio
        }
    
    def generar(
        self,
        docente: DocenteData,
        payment: PaymentData,
        descripcion_completa: str,
        carpeta_docente: str
    ) -> str:
        # Obtener plantilla según tipo de contrato
        tipo_contrato = "contrato" if docente.es_contrato else "tercero"
        ruta_template = self.doc_service.obtener_ruta_plantilla(
            'conformidad', 
            tipo_contrato, 
            self.config
        )
        
        if not PathUtils.archivo_existe(ruta_template):
            raise FileNotFoundError(f"No se encontró la plantilla: {ruta_template}")
        
        # Construir reemplazos
        reemplazos = self.construir_reemplazos(docente, payment, descripcion_completa)
        
        # Generar nombre de archivo
        nombre_archivo = f"CONFORMIDAD - {docente.nombre_limpio} - {self.config.mes} {self.config.anio}.docx"
        ruta_salida = os.path.join(carpeta_docente, nombre_archivo)
        
        # Crear directorio si no existe
        PathUtils.crear_directorio(carpeta_docente)
        
        # Generar documento
        self.doc_service.generar_documento(ruta_template, reemplazos, ruta_salida)
        
        # Limpiar texto administrativo si aplica
        if self.config.es_administrativo:
            self.doc_service.eliminar_texto_administrativo(ruta_salida)
        
        return ruta_salida
