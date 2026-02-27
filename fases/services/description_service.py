from typing import List
from fases.models import PaymentData, CursoDetalle


class DescriptionService:
    
    @staticmethod
    def redactar_cursos(cadena_cursos: str, tiene_bono: bool = False, tiene_servicio_actualizacion: bool = False) -> str:
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
        
        # Agregar servicio de actualización si existe y no está ya en la lista
        if tiene_servicio_actualizacion:
            # Verificar que no esté ya agregado desde la columna Curso
            if not any("Servicio de actualización" in elem for elem in elementos_descripcion):
                elementos_descripcion.append("Servicio de actualización de materiales de enseñanza")
        
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
            "• Dictar clases, preparar las clases, evaluar a los alumnos, diseñar exámenes, entregar notas y presentar informe de dictado de curso."
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
    
    # ============= NUEVOS MÉTODOS CON MODALIDADES ============= 
    
    @staticmethod
    def redactar_servicios_con_modalidad(cursos_detallados: List[CursoDetalle]) -> str:
        """
        Genera descripción de servicios con modalidad específica para cada uno.
        Usado en OFICIO, COTIZACIÓN y CONFORMIDAD.
        
        Ejemplo:
        "28 horas de clases de Inglés Avanzado 2 bajo la modalidad presencial,
        28 horas de clases de Inglés Avanzado 9 bajo la modalidad virtual,
        10 horas de examen de clasificación bajo la modalidad virtual
        y 8 horas de diseño de exámenes"
        
        Args:
            cursos_detallados: Lista de objetos CursoDetalle
        
        Returns:
            Descripción completa con modalidades
        """
        if not cursos_detallados:
            return "N/A"
        
        descripciones = []
        for curso in cursos_detallados:
            descripciones.append(curso.generar_descripcion_individual())
        
        # Unir con el formato adecuado
        if len(descripciones) == 1:
            # Un solo elemento
            if cursos_detallados[0].es_servicio_especial or cursos_detallados[0].tipo_servicio == "BONO":
                return descripciones[0]
            else:
                return f"servicio de dictado de {descripciones[0]}"
        
        elif len(descripciones) == 2:
            # Dos elementos, usar "y"
            primer_elem = descripciones[0]
            segundo_elem = descripciones[1]
            
            # Si el primero es curso, agregar prefijo
            if cursos_detallados[0].es_curso_academico:
                primer_elem = f"servicio de dictado de {primer_elem}"
            
            return f"{primer_elem} y {segundo_elem}"
        
        else:
            # Más de dos elementos, usar comas y "y" al final
            primer_elem = descripciones[0]
            if cursos_detallados[0].es_curso_academico:
                primer_elem = f"servicio de dictado de {primer_elem}"
            
            # Todos los elementos del medio con comas
            elementos_medio = descripciones[1:-1]
            ultimo_elem = descripciones[-1]
            
            partes = [primer_elem] + elementos_medio
            return f"{', '.join(partes)} y {ultimo_elem}"
    
    @staticmethod
    def agrupar_por_modalidad_tdr(cursos_detallados: List[CursoDetalle]) -> str:
        """
        Agrupa cursos/servicios por modalidad para el TDR.
        
        Formato de salida:
        "Presencial: Inglés Avanzado 2
        Virtual: Inglés Avanzado 9, Examen de clasificación"
        
        Args:
            cursos_detallados: Lista de objetos CursoDetalle
        
        Returns:
            Texto agrupado por modalidad
        """
        if not cursos_detallados:
            return "N/A"
        
        # Agrupar por modalidad
        presencial = []
        virtual = []
        otras = []
        
        for curso in cursos_detallados:
            # Excluir servicios sin modalidad (diseño de exámenes, bono)
            if curso.es_sin_modalidad:
                continue
            
            modalidad = curso.modalidad_texto.lower()
            
            if modalidad == "presencial":
                presencial.append(curso.nombre)
            elif modalidad == "virtual":
                virtual.append(curso.nombre)
            else:
                otras.append(f"{curso.nombre} ({modalidad})")
        
        # Construir texto final
        lineas = []
        
        if presencial:
            lineas.append(f"Presencial: {', '.join(presencial)}")
        
        if virtual:
            lineas.append(f"Virtual: {', '.join(virtual)}")
        
        if otras:
            lineas.append(f"Otras: {', '.join(otras)}")
        
        return "\n".join(lineas) if lineas else "N/A"
