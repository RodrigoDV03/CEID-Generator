import os
from typing import Dict
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import PaymentService, DocumentGeneratorService
from fases.utils import PathUtils


class ControlAvanceBuilder:
    
    def __init__(self, config: DocumentConfig):
        self.config = config
        self.payment_service = PaymentService()
        self.doc_service = DocumentGeneratorService()
    
    def construir_reemplazos(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> Dict[str, str]:
        # Calcular todos los saldos
        saldos = self.payment_service.calcular_saldos_armadas(payment)
        
        # Asegurar que el número de contrato no esté vacío
        numero_contrato = docente.numero_contrato if docente.numero_contrato else "001"
        
        return {
            "Nombre_Docente": docente.nombre,
            "Nro_Contrato": numero_contrato,
            "Idioma_Docente": docente.especialidad,
            "Monto_Total": saldos['Monto_Total'],
            "Total_Primera": saldos['Total_Primera'],
            "Total_Segunda": saldos['Total_Segunda'],
            "Total_Tercera": saldos['Total_Tercera'],
            "Saldo_Restante": saldos['Saldo_Restante'],
            "Saldo_Primera": saldos['Saldo_Primera'],
            "Saldo_Segunda": saldos['Saldo_Segunda']
        }
    
    def generar(
        self,
        docente: DocenteData,
        payment: PaymentData,
        carpeta_docente: str
    ) -> str:
        # Validar que se especificó número de armada
        if self.config.numero_armada == 'sin armada':
            raise ValueError("No se puede generar control de avance sin número de armada")
        
        # Obtener plantilla según número de armada
        ruta_template = self.doc_service.obtener_ruta_plantilla(
            'control', 
            '', 
            self.config
        )
        
        if not PathUtils.archivo_existe(ruta_template):
            raise FileNotFoundError(f"No se encontró la plantilla: {ruta_template}")
        
        # Construir reemplazos
        reemplazos = self.construir_reemplazos(docente, payment)
        
        # Generar nombre de archivo
        nombre_archivo = f"CONTROL DE AVANCE - {docente.nombre_limpio} - {self.config.mes} {self.config.anio}.docx"
        ruta_salida = os.path.join(carpeta_docente, nombre_archivo)
        
        # Crear directorio si no existe
        PathUtils.crear_directorio(carpeta_docente)
        
        # Generar documento
        self.doc_service.generar_documento(ruta_template, reemplazos, ruta_salida)
        
        return ruta_salida
