import datetime
import pandas as pd
from typing import Optional
from fases.models import DocenteData, PaymentData, DocumentConfig
from fases.services import (
    ExcelReaderService, 
    DocenteService, 
    PaymentService, 
    DescriptionService
)
from .builders import ConformidadBuilder, ControlAvanceBuilder


class FaseFinalGenerator:
    
    def __init__(self, config: DocumentConfig):
        self.config = config
        
        # Servicios
        self.excel_service = ExcelReaderService()
        self.docente_service = DocenteService()
        self.payment_service = PaymentService()
        self.description_service = DescriptionService()
        
        # Builders
        self.conformidad_builder = ConformidadBuilder(config)
        self.control_builder = ControlAvanceBuilder(config)
    
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
            tiene_bono=payment.tiene_bono
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
    
    def generar_conformidad(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> str:
        # Crear carpeta del docente
        carpeta_docente = self.docente_service.crear_carpeta_docente(
            self.config.carpeta_destino,
            self.config,
            docente
        )
        
        # Generar descripción completa
        descripcion_completa = self.generar_descripcion_completa(docente, payment)
        
        # Generar documento
        ruta = self.conformidad_builder.generar(
            docente, 
            payment, 
            descripcion_completa, 
            carpeta_docente
        )
        
        return ruta
    
    def generar_control_avance(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> Optional[str]:
        # Solo para docentes con contrato
        if not docente.es_contrato:
            return None
        
        # Solo si se especificó número de armada
        if self.config.numero_armada == 'sin armada':
            return None
        
        # Crear carpeta del docente
        carpeta_docente = self.docente_service.crear_carpeta_docente(
            self.config.carpeta_destino,
            self.config,
            docente
        )
        
        # Generar documento
        ruta = self.control_builder.generar(
            docente,
            payment,
            carpeta_docente
        )
        
        return ruta
    
    def procesar_docente(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> None:
        # Validar docente
        es_valido, mensaje = self.docente_service.validar_docente(docente)
        if not es_valido:
            print(f"❌ Error - {mensaje}")
            return
        
        try:
            # Generar conformidad (para todos)
            self.generar_conformidad(docente, payment)
            print(f"✅ {docente.nombre} - Documento de conformidad generado correctamente.")
            
        except Exception as e:
            print(f"❌ Error generando conformidad para {docente.nombre}: {e}")
    
    def procesar_control_pagos(
        self,
        df_control: 'pd.DataFrame'
    ) -> None:
        for _, fila in df_control.iterrows():
            try:
                # Extraer nombre del docente
                nombre = self.excel_service.extraer_docente_nombre_control(fila)
                if nombre == "N/A":
                    continue
                
                # Crear objeto DocenteData básico
                docente = DocenteData(
                    nombre=nombre,
                    dni="",
                    ruc="",
                    especialidad=str(getattr(fila, "Especialidad", "")),
                    numero_contrato=self.excel_service.extraer_numero_contrato_control(fila),
                    estado_docente="CONTRATO"
                )
                
                # Extraer datos de pago del control
                payment = self.excel_service.extraer_payment_data_control(fila)
                
                # Generar control de avance
                ruta = self.generar_control_avance(docente, payment)
                if ruta:
                    print(f"✅ {nombre} - Control de pagos generado correctamente.")
                
            except Exception as e:
                print(f"❌ Error procesando control para {nombre}: {e}")
                continue


def procesar_planilla_fase_final(
    planilla_path: str,
    excel_control_pagos: Optional[str],
    hoja: str,
    carpeta_salida: str,
    mes: str,
    numero_armada: str,
    tipo_fase_final: str
) -> None:
    # Crear configuración
    config = DocumentConfig(
        mes=mes,
        anio=datetime.datetime.now().year,
        numero_armada=numero_armada,
        tipo_fase="final",
        tipo_docente=tipo_fase_final,
        carpeta_destino=carpeta_salida
    )
    
    # Crear generador
    generador = FaseFinalGenerator(config)
    
    # Leer planilla principal
    excel_service = ExcelReaderService()
    df = excel_service.leer_planilla(planilla_path, hoja)
    
    # Procesar cada docente (conformidad)
    for _, fila in df.iterrows():
        try:
            # Extraer datos
            docente = excel_service.extraer_docente_data(fila)
            payment = excel_service.extraer_payment_data(fila)
            
            # Procesar docente
            generador.procesar_docente(docente, payment)
            
        except ValueError as e:
            print(f"Error: {e}")
            continue
        except Exception as e:
            print(f"Error inesperado: {e}")
            continue
    
    # Procesar control de pagos si aplica
    if tipo_fase_final == "planilla docente (con contrato)" and excel_control_pagos:
        try:
            df_control = excel_service.leer_control_pagos(excel_control_pagos)
            generador.procesar_control_pagos(df_control)
        except Exception as e:
            print(f"Error procesando control de pagos: {e}")
