from abc import ABC, abstractmethod
from typing import Optional
from .config import TipoCorreo, EmailConfig, AÑO_ACTUAL


class EmailBuilder(ABC):
    
    def __init__(self):
        """Inicializa el builder con valores por defecto."""
        self._mes: Optional[str] = None
        self._anio: int = AÑO_ACTUAL
        self._firma_html: str = ""
        self._nombre: Optional[str] = None
    
    def con_mes(self, mes: str) -> 'EmailBuilder':
        self._mes = mes
        return self
    
    def con_anio(self, anio: int) -> 'EmailBuilder':
        self._anio = anio
        return self
    
    def con_firma(self, firma_html: str) -> 'EmailBuilder':
        self._firma_html = firma_html
        return self
    
    def con_nombre(self, nombre: str) -> 'EmailBuilder':
        self._nombre = nombre
        return self
    
    @abstractmethod
    def construir_cuerpo(self) -> str:
        pass
    
    @abstractmethod
    def construir_asunto(self) -> str:
        pass
    
    def _validar_datos(self) -> None:
        if not self._mes:
            raise ValueError("Mes es requerido para construir el correo")
    
    def _crear_encabezado_html(self) -> str:

        return f"""
<html>
    <body style="font-family: {EmailConfig.FUENTE}; font-size: {EmailConfig.TAMAÑO_FUENTE};">
        <p>Buen día:</p>
"""
    
    def _crear_pie_html(self) -> str:
        return f"""
        {self._firma_html}
    </body>
</html>
"""
    
    def _crear_texto_destacado(self, texto: str) -> str:
        return f'<span style="font-weight: bold; color: {EmailConfig.COLOR_DESTACADO};">{texto}</span>'


class EmailDocenteBuilder(EmailBuilder):
    
    def __init__(self):
        super().__init__()
        self._servicio: Optional[str] = None
    
    def con_servicio(self, servicio: str) -> 'EmailDocenteBuilder':
        self._servicio = servicio
        return self
    
    def _validar_datos(self) -> None:
        super()._validar_datos()
        if not self._servicio:
            raise ValueError("Servicio es requerido para correos de docentes")
        if not self._nombre:
            raise ValueError("Nombre es requerido para construir el asunto")
    
    def construir_asunto(self) -> str:
        self._validar_datos()
        
        # Extraer solo los apellidos (antes de la coma)
        nombre_formato = self._nombre.split(",")[0].strip()
        
        return (
            f"Envío de orden de servicio y solicitud de recibo por honorarios – "
            f"{self._mes} {self._anio} - {nombre_formato}"
        )
    
    def construir_cuerpo(self) -> str:
        self._validar_datos()
        
        concepto_destacado = self._crear_texto_destacado(
            f"Servicio de dictado de {self._servicio}, BAJO LA MODALIDAD {EmailConfig.MODALIDAD}."
        )
        
        plazo_destacado = self._crear_texto_destacado(
            f"{EmailConfig.DIAS_VENCIMIENTO} días"
        )
        
        formato_destacado = self._crear_texto_destacado(
            f"formato {EmailConfig.FORMATO_RECIBO}"
        )
        
        cuerpo_contenido = f"""
        <p>
            Adjunto su orden de servicio correspondiente al mes de {self._mes} {self._anio}. 
            Con este documento, ya puede proceder con la emisión de su recibo por honorarios. 
            Para evitar retrasos en el pago, tenga en cuenta lo siguiente:
        </p>

        <ul>
            <li>
                El concepto del recibo por honorarios es: {concepto_destacado}
            </li>

            <li>
                El pago se realizará a crédito, con un plazo de vencimiento de {plazo_destacado} 
                desde la fecha de emisión del recibo en la plataforma SUNAT.
            </li>
        </ul>

        <p>
            Una vez emitido, envíe el recibo en {formato_destacado} como respuesta a este mismo correo, 
            sin generar un nuevo hilo.
        </p>
"""
        
        return self._crear_encabezado_html() + cuerpo_contenido + self._crear_pie_html()


class EmailAdministrativoBuilder(EmailBuilder):
    def construir_asunto(self) -> str:
        self._validar_datos()
        
        return (
            f"Envío de orden de servicio y solicitud de recibo por honorarios – "
            f"{self._mes} {self._anio}"
        )
    
    def construir_cuerpo(self) -> str:
        self._validar_datos()
        
        formato_destacado = self._crear_texto_destacado(
            f"formato {EmailConfig.FORMATO_RECIBO}"
        )
        
        cuerpo_contenido = f"""
        <p>
            Adjunto su orden de servicio correspondiente al mes de {self._mes} {self._anio}. 
            Con este documento, ya puede proceder con la emisión de su recibo por honorarios.
        </p>

        <p>
            Una vez emitido, envíe el recibo en {formato_destacado} como respuesta a este mismo correo, 
            sin generar un nuevo hilo.
        </p>
"""
        
        return self._crear_encabezado_html() + cuerpo_contenido + self._crear_pie_html()


class EmailBuilderFactory:
    @staticmethod
    def crear_builder(tipo: TipoCorreo) -> EmailBuilder:
        if tipo == TipoCorreo.DOCENTE:
            return EmailDocenteBuilder()
        elif tipo == TipoCorreo.ADMINISTRATIVO:
            return EmailAdministrativoBuilder()
        else:
            raise ValueError(f"Tipo de correo no válido: {tipo}")


# Funciones de compatibilidad con código legacy
def generar_cuerpo_correo_docente_html(
    mes: str, 
    anio: str, 
    servicio: str, 
    firma_html: str = ""
) -> str:
    builder = EmailDocenteBuilder()
    return (builder
            .con_mes(mes)
            .con_anio(int(anio))
            .con_servicio(servicio)
            .con_firma(firma_html)
            .construir_cuerpo())


def generar_cuerpo_correo_administrativo_html(
    mes: str, 
    anio: str, 
    firma_html: str = ""
) -> str:
    builder = EmailAdministrativoBuilder()
    return (builder
            .con_mes(mes)
            .con_anio(int(anio))
            .con_firma(firma_html)
            .construir_cuerpo())
