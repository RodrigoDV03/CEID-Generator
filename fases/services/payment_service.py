from typing import Dict, Tuple
from fases.models import PaymentData


class PaymentService:
    
    @staticmethod
    def calcular_montos_completos(payment: PaymentData, es_administrativo: bool) -> Dict[str, any]:
        # Determinar qué bono mostrar según el tipo
        bono_para_mostrar = 0 if es_administrativo else payment.bono
        
        # Calcular montos principales
        monto_total = payment.calcular_monto_total(es_administrativo)
        monto_sin_actualizacion = payment.calcular_monto_sin_actualizacion(es_administrativo)
        
        return {
            'monto_total': monto_total,
            'monto_sin_actualizacion': monto_sin_actualizacion,
            'servicio_actualizacion': payment.servicio_actualizacion,
            'bono_para_mostrar': bono_para_mostrar,
            'categoria_monto': payment.categoria_monto,
            
            # Versiones en letras
            'monto_total_letras': payment.monto_a_letras(monto_total),
            'monto_sin_actualizacion_letras': payment.monto_a_letras(monto_sin_actualizacion),
            'servicio_actualizacion_letras': payment.monto_a_letras(payment.servicio_actualizacion),
            'bono_letras': payment.monto_a_letras(bono_para_mostrar),
            'categoria_letras': payment.monto_a_letras(payment.categoria_monto),
            
            # Versiones formateadas
            'monto_total_formato': payment.formatear_con_letras(monto_total),
            'categoria_formato': payment.formatear_con_letras(payment.categoria_monto)
        }
    
    @staticmethod
    def generar_descripcion_horas(payment: PaymentData) -> Tuple[str, str]:
        horas_disenio_num = payment.calcular_horas_disenio()
        horas_clasif_num = payment.calcular_horas_clasificacion()
        
        horas_disenio = f"{horas_disenio_num} horas de diseño de exámenes"
        
        if horas_clasif_num == 1:
            horas_clasif = f"{horas_clasif_num} hora de examen de clasificación"
        else:
            horas_clasif = f"{horas_clasif_num} horas de examen de clasificación"
        
        return horas_disenio, horas_clasif
    
    @staticmethod
    def generar_monto_referencial(
        monto_sin_actualizacion: float,
        monto_sin_actualizacion_letras: str,
        servicio_actualizacion: float,
        servicio_actualizacion_letras: str,
        bono: float,
        bono_letras: str,
        monto_total: float,
        monto_total_letras: str
    ) -> str:
        lineas = []
        
        # Agregar monto base si existe
        if monto_sin_actualizacion > 0:
            lineas.append(f"S/. {monto_sin_actualizacion:,.2f} ({monto_sin_actualizacion_letras}).")
        
        # Agregar servicio de actualización si existe
        if servicio_actualizacion > 0:
            lineas.append(
                f"S/. {servicio_actualizacion:,.2f} ({servicio_actualizacion_letras}) "
                "por servicio de actualización de materiales de enseñanza."
            )
        
        # Agregar bono si existe
        if bono > 0:
            lineas.append(
                f"S/. {bono:,.2f} ({bono_letras}) "
                "por servicio de diseño y evaluación del examen anual."
            )
        
        # Si hay más de un concepto, agregar el monto total
        if len(lineas) > 1:
            lineas.append(
                f"Monto total: S/. {monto_total:,.2f} ({monto_total_letras}). "
                "Incluye el impuesto y la contribución de ley"
            )
        else:
            # Si solo hay un concepto, solo agregar la nota de impuestos
            lineas.append("Incluye el impuesto y la contribución de ley")
            lineas_texto = lineas[0]
            if not lineas_texto.endswith("."):
                lineas_texto += "."
            return lineas_texto.replace(".", ". Incluye el impuesto y la contribución de ley", 1)
        
        return "\n".join(lineas)
    
    @staticmethod
    def calcular_saldos_armadas(payment: PaymentData) -> Dict[str, str]:
        return {
            'Monto_Total': payment.formatear_monto(payment.monto_total_contrato),
            'Total_Primera': payment.formatear_monto(payment.primera_armada),
            'Total_Segunda': payment.formatear_monto(payment.segunda_armada),
            'Total_Tercera': payment.formatear_monto(payment.tercera_armada),
            'Saldo_Restante': payment.formatear_monto(payment.calcular_saldo_restante()),
            'Saldo_Primera': payment.formatear_monto(payment.calcular_saldo_primera()),
            'Saldo_Segunda': payment.formatear_monto(payment.calcular_saldo_segunda())
        }
