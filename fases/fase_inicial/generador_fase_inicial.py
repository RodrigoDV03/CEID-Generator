import datetime
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import (
    ExcelReaderService, 
    DocenteService, 
    PaymentService, 
    DescriptionService
)
from .builders import OficioBuilder, TdrBuilder, CotizacionBuilder


class FaseInicialGenerator:
    
    def __init__(self, config: DocumentConfig):
        self.config = config
        
        # Servicios
        self.excel_service = ExcelReaderService()
        self.docente_service = DocenteService()
        self.payment_service = PaymentService()
        self.description_service = DescriptionService()
        
        # Builders
        self.oficio_builder = OficioBuilder(config)
        self.tdr_builder = TdrBuilder(config)
        self.cotizacion_builder = CotizacionBuilder(config)
    
    def generar_descripcion_completa(
        self, 
        docente: DocenteData, 
        payment: PaymentData
    ) -> str:
        # Para administrativos, retornar el curso directamente
        if self.config.es_administrativo:
            return docente.curso
        
        # Para docentes, redactar cursos con formato
        descripcion_base = self.description_service.redactar_cursos(
            docente.curso,
            tiene_bono=payment.tiene_bono,
            tiene_servicio_actualizacion=payment.tiene_servicio_actualizacion
        )
        
        # Generar descripciones de horas
        horas_disenio, horas_clasif = self.payment_service.generar_descripcion_horas(payment)
        
        # Generar descripción completa
        return self.description_service.generar_descripcion_completa(
            descripcion_base,
            payment,
            horas_disenio,
            horas_clasif,
            es_administrativo=self.config.es_administrativo
        )
    
    def generar_documentos_docente(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> None:
        # Validar docente
        es_valido, mensaje = self.docente_service.validar_docente(docente)
        if not es_valido:
            print(f"❌ Error - {mensaje}")
            return
        
        # Crear carpeta del docente
        carpeta_docente = self.docente_service.crear_carpeta_docente(
            self.config.carpeta_destino,
            self.config,
            docente
        )
        
        # Generar descripción completa
        descripcion_completa = self.generar_descripcion_completa(docente, payment)
        
        try:
            # 1. Generar OFICIO (para todos)
            self.oficio_builder.generar(docente, payment, descripcion_completa, carpeta_docente)
            
            # 2. Generar TDR (solo para terceros)
            if docente.es_tercero:
                self.tdr_builder.generar(docente, payment, descripcion_completa, carpeta_docente)
            
            # 3. Generar COTIZACIÓN (para administrativos o terceros)
            if self.config.es_administrativo or docente.es_tercero:
                self.cotizacion_builder.generar(docente, payment, descripcion_completa, carpeta_docente)
            
            print(f"✅ {docente.nombre} - Documentos generados correctamente.")
            
        except Exception as e:
            print(f"❌ Error generando documentos para {docente.nombre}: {e}")


def procesar_planilla_fase_inicial(
    planilla_path: str,
    hoja_seleccionada: str,
    carpeta_destino: str,
    mes: str,
    numero_armada: str,
    tipo_fase_inicial: str
) -> None:
    # Crear configuración
    config = DocumentConfig(
        mes=mes,
        anio=datetime.datetime.now().year,
        numero_armada=numero_armada,
        tipo_fase="inicial",
        tipo_docente=tipo_fase_inicial,
        carpeta_destino=carpeta_destino
    )
    
    # Crear generador
    generador = FaseInicialGenerator(config)
    
    # Leer planilla
    excel_service = ExcelReaderService()
    df = excel_service.leer_planilla(planilla_path, hoja_seleccionada)
    
    # Procesar cada fila
    for i, fila in enumerate(df.itertuples(index=False), start=1):
        try:
            # Extraer datos
            docente = excel_service.extraer_docente_data(fila)
            payment = excel_service.extraer_payment_data(fila)
            
            # Generar documentos
            generador.generar_documentos_docente(docente, payment)
            
        except ValueError as e:
            print(f"Fila {i}: {e}")
            continue
        except Exception as e:
            print(f"Fila {i}: Error inesperado - {e}")
            continue
