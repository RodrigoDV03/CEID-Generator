import os
from typing import Dict
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import PaymentService, DescriptionService, DocumentGeneratorService


class OficioBuilder:
    
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
        
        # Debug: mostrar el número de contrato
        if not docente.numero_contrato:
            print(f"⚠️  {docente.nombre} - Número de contrato vacío, usando valor por defecto: {numero_contrato}")
        
        return {
            "Nro_Contrato": numero_contrato,
            "docente": docente.nombre,
            "descripcion": descripcion_completa,
            "categoria": montos['categoria_formato'],
            "monto_subtotal": montos['monto_total_formato'],
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
            'oficio', 
            tipo_contrato, 
            self.config
        )
        
        # Construir reemplazos
        reemplazos = self.construir_reemplazos(docente, payment, descripcion_completa)
        
        # Generar nombre de archivo
        nombre_archivo = f"OFICIO - {docente.nombre_limpio} - {self.config.mes} {self.config.anio}.docx"
        ruta_salida = os.path.join(carpeta_docente, nombre_archivo)
        
        # Generar documento
        self.doc_service.generar_documento(ruta_template, reemplazos, ruta_salida)
        
        # Limpiar texto administrativo si aplica
        if self.config.es_administrativo:
            self.doc_service.eliminar_texto_administrativo(ruta_salida)
        
        return ruta_salida
