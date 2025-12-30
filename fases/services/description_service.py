from typing import List
from fases.models import PaymentData


class DescriptionService:
    
    @staticmethod
    def redactar_cursos(cadena_cursos: str, tiene_bono: bool = False) -> str:
        if not isinstance(cadena_cursos, str):
            return "N/A"
        
        cursos = [c.strip() for c in cadena_cursos.split("/") if c.strip()]
        if not cursos:
            return "N/A"
        
        elementos_descripcion = []
        
        for curso in cursos:
            # Detectar si es el servicio de actualización de materiales
            if "Servicio de actualización de materiales de enseñanza" in curso:
                elementos_descripcion.append(curso.strip())
            else:
                # Para cursos académicos normales, agregar el formato tradicional
                elementos_descripcion.append(f"28 horas de clases de {curso}")
        
        # Agregar el bono al final si existe
        if tiene_bono:
            elementos_descripcion.append("servicio de diseño y evaluación del examen anual")
        
        return DescriptionService._unir_elementos_descripcion(elementos_descripcion)
    
    @staticmethod
    def _unir_elementos_descripcion(elementos: List[str]) -> str:
        if len(elementos) == 1:
            # Si solo hay un elemento y es servicio especial, retornarlo directo
            if any(
                keyword in elementos[0] 
                for keyword in [
                    "Servicio de actualización de materiales de enseñanza",
                    "servicio de diseño y evaluación del examen anual"
                ]
            ):
                return elementos[0]
            else:
                return f"servicio de dictado de {elementos[0]}"
        
        elif len(elementos) == 2:
            # Si hay servicio especial como segundo elemento, usar coma
            if any(
                keyword in elementos[1]
                for keyword in [
                    "Servicio de actualización de materiales de enseñanza",
                    "servicio de diseño y evaluación del examen anual"
                ]
            ):
                return f"servicio de dictado de {elementos[0]}, {elementos[1].lower()}"
            else:
                return f"servicio de dictado de {elementos[0]}, {elementos[1]}"
        else:
            # Si hay más de 2 elementos, usar solo comas
            return f"servicio de dictado de {', '.join(elementos)}"
    
    @staticmethod
    def generar_descripcion_completa(
        descripcion_base: str,
        payment: PaymentData,
        horas_disenio: str,
        horas_clasificacion: str,
        es_administrativo: bool = False
    ) -> str:
        if es_administrativo:
            return descripcion_base
        
        # Sin horas adicionales
        if not payment.tiene_disenio_examenes and not payment.tiene_examen_clasificacion:
            return descripcion_base
        
        # Solo diseño
        if payment.tiene_disenio_examenes and not payment.tiene_examen_clasificacion:
            return f"{descripcion_base} y {horas_disenio}"
        
        # Solo clasificación (raro pero posible)
        if not payment.tiene_disenio_examenes and payment.tiene_examen_clasificacion:
            return f"{descripcion_base} y {horas_clasificacion}"
        
        # Ambos
        return f"{descripcion_base}, {horas_disenio} y {horas_clasificacion}"
    
    @staticmethod
    def generar_actividades_docentes(payment: PaymentData) -> str:
        actividades_base = (
            "• Dictar clases, preparar las clases, evaluar a los alumnos, "
            "diseñar exámenes, entregar notas y presentar informe de dictado de curso."
        )
        
        actividades_actualizacion = (
            "• Revisar y actualizar los materiales de enseñanza conforme a los planes de estudio vigentes.\n"
            "• Elaborar y mejorar materiales didácticos y recursos pedagógicos para clases presenciales y virtuales.\n"
            "• Actualizar instrumentos de evaluación (exámenes y prácticas)."
        )
        
        # Solo actualización (sin clases)
        if payment.tiene_servicio_actualizacion and payment.calcular_monto_sin_actualizacion(False) == 0:
            return actividades_actualizacion
        
        # Clases + actualización
        if payment.tiene_servicio_actualizacion:
            return f"{actividades_base}\n{actividades_actualizacion}"
        
        # Solo clases (caso normal)
        return actividades_base
    
    @staticmethod
    def generar_actividades_admin_cotizacion(payment: PaymentData) -> str:
        actividades_base = """-	Dictar clases.
-	Preparar las clases.
-	Evaluar a los alumnos.
-	Diseñar exámenes.
-	Entregar acta de nota.
-	Presentar informe de dictado de curso."""
        
        actividades_actualizacion = """-	Revisar y actualizar los materiales de enseñanza conforme a los planes de estudio vigentes.
-	Elaborar y mejorar materiales didácticos y recursos pedagógicos para clases presenciales y virtuales.
-	Actualizar instrumentos de evaluación (exámenes y prácticas)."""
        
        monto_sin_actualizacion = payment.calcular_monto_sin_actualizacion(False)
        
        # Solo actualización
        if payment.tiene_servicio_actualizacion and monto_sin_actualizacion == 0:
            return actividades_actualizacion
        
        # Clases + actualización
        if payment.tiene_servicio_actualizacion and monto_sin_actualizacion > 0:
            return f"{actividades_base}\n{actividades_actualizacion}"
        
        # Solo clases
        return actividades_base
