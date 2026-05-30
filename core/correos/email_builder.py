from abc import ABC, abstractmethod
import re
from html import escape
from typing import Optional
from .config import TipoCorreo, EmailConfig, AÑO_ACTUAL, normalizar_servicio_para_correo


class EmailBuilder(ABC):
    
    def __init__(self):
        """Inicializa el builder con valores por defecto."""
        self._mes: Optional[str] = None
        self._anio: int = AÑO_ACTUAL
        self._firma_html: str = ""
        self._nombre: Optional[str] = None
        self._mes_inicio_contrato: Optional[str] = None
        self._mes_fin_contrato: Optional[str] = None
        self._es_reconocimiento_deuda: bool = False
    
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

    def con_periodo_contrato(self, mes_inicio: str, mes_fin: str) -> 'EmailBuilder':
        self._mes_inicio_contrato = mes_inicio
        self._mes_fin_contrato = mes_fin
        return self

    def con_reconocimiento_deuda(self, es_reconocimiento_deuda: bool) -> 'EmailBuilder':
        self._es_reconocimiento_deuda = es_reconocimiento_deuda
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

    def _formatear_texto_marcado(self, texto: str) -> str:
        partes = re.split(r'(\*[^*]+\*)', texto)
        salida = []

        for parte in partes:
            if not parte:
                continue
            if parte.startswith('*') and parte.endswith('*'):
                contenido = parte[1:-1].strip()
                salida.append(self._crear_texto_destacado(escape(contenido)))
            else:
                salida.append(escape(parte))

        return ''.join(salida)

    def _crear_parrafo_html(self, texto: str) -> str:
        return f'<p>{self._formatear_texto_marcado(texto)}</p>'

    def _crear_lista_html(self, items: list[str]) -> str:
        items_html = ''.join(f'<li>{self._formatear_texto_marcado(item)}</li>' for item in items)
        return f'<ul style="margin: 0 0 0 22px; padding-left: 18px;">{items_html}</ul>'


class EmailDocenteBuilder(EmailBuilder):
    
    def __init__(self):
        super().__init__()
        self._servicio: Optional[str] = None
        self._modalidad: Optional[str] = None
    
    def con_servicio(self, servicio: str) -> 'EmailDocenteBuilder':
        self._servicio = servicio
        return self

    def con_modalidad(self, modalidad: str) -> 'EmailDocenteBuilder':
        self._modalidad = modalidad
        return self
    
    def _validar_datos(self) -> None:
        super()._validar_datos()
        if not self._servicio:
            raise ValueError("Servicio es requerido para correos de docentes")
        if not self._modalidad:
            raise ValueError("Modalidad es requerida para correos de docentes")
        if not self._nombre:
            raise ValueError("Nombre es requerido para construir el asunto")

    def _construir_concepto_servicio(self) -> str:
        return normalizar_servicio_para_correo(self._servicio)
    
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

        concepto_servicio = self._construir_concepto_servicio()

        texto_orden = (
            f"correspondiente al reconocimiento de deuda de *{self._mes} {self._anio}*"
            if self._es_reconocimiento_deuda
            else f"correspondiente al mes de *{self._mes} {self._anio}*."
        )

        if not texto_orden.endswith('.'):
            texto_orden += '.'

        cuerpo_contenido = (
            self._crear_parrafo_html(
                f"Adjunto su orden de servicio {texto_orden} Con este documento, ya puede proceder con la emisión de su recibo por honorarios. Para evitar retrasos en el pago, tenga en cuenta lo siguiente:"
            )
            + self._crear_lista_html([
                f"Número de RUC de entidad usuaria: *20148092282* (Número de RUC de la UNMSM).",
                f"Concepto del recibo: *{concepto_servicio}.*",
                "Pago: *al crédito.*",
                f"Plazo de vencimiento: *{EmailConfig.DIAS_VENCIMIENTO} días desde la emisión del recibo* en la plataforma de SUNAT.",
            ])
            + self._crear_parrafo_html(
                f"Una vez emitido, envíe el recibo en formato *{EmailConfig.FORMATO_RECIBO}* como respuesta a este mismo correo, sin generar un nuevo hilo."
            )
        )

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

        texto_orden = (
            f"correspondiente al reconocimiento de deuda de *{self._mes} {self._anio}*"
            if self._es_reconocimiento_deuda
            else f"correspondiente al mes de *{self._mes} {self._anio}*."
        )

        if not texto_orden.endswith('.'):
            texto_orden += '.'

        cuerpo_contenido = (
            self._crear_parrafo_html(
                f"Adjunto su orden de servicio {texto_orden} Con este documento, ya puede proceder con la emisión de su recibo por honorarios. Para evitar retrasos en el pago, tenga en cuenta lo siguiente:"
            )
            + self._crear_lista_html([
                f"Número de RUC de entidad usuaria: *20148092282* (Número de RUC de la UNMSM).",
                "Pago: *al crédito.*",
                f"Plazo de vencimiento: *{EmailConfig.DIAS_VENCIMIENTO} días desde la emisión del recibo* en la plataforma de SUNAT.",
            ])
            + self._crear_parrafo_html(
                f"Una vez emitido, envíe el recibo en formato *{EmailConfig.FORMATO_RECIBO}* como respuesta a este mismo correo, sin generar un nuevo hilo."
            )
        )

        return self._crear_encabezado_html() + cuerpo_contenido + self._crear_pie_html()


class EmailDocenteContratoBuilder(EmailDocenteBuilder):
    def _validar_datos(self) -> None:
        super()._validar_datos()
        if not self._mes_inicio_contrato or not self._mes_fin_contrato:
            raise ValueError("Periodo de contrato es requerido para primera vez")

    def construir_asunto(self) -> str:
        asunto_base = super().construir_asunto()
        return f"{asunto_base} | Contrato de locacion de servicios"

    def construir_cuerpo(self) -> str:
        self._validar_datos()

        concepto_servicio = self._construir_concepto_servicio()

        texto_orden = (
            f"correspondiente al reconocimiento de deuda de *{self._mes} {self._anio}*"
            if self._es_reconocimiento_deuda
            else f"correspondiente al mes de *{self._mes} {self._anio}*.")

        if not texto_orden.endswith('.'):
            texto_orden += '.'

        cuerpo_contenido = (
            self._crear_parrafo_html(
                f"Adjunto su orden de servicio {texto_orden} Con este documento, ya puede proceder con la emisión de su recibo por honorarios. Para evitar retrasos en el pago, tenga en cuenta lo siguiente:"
            )
            + self._crear_lista_html([
                f"Número de RUC de entidad usuaria: *20148092282* (Número de RUC de la UNMSM).",
                f"Concepto del recibo: *{concepto_servicio}.*",
                "Pago: *al crédito.*",
                f"Plazo de vencimiento: *{EmailConfig.DIAS_VENCIMIENTO} días desde la emisión del recibo* en la plataforma de SUNAT.",
            ])
            + self._crear_parrafo_html(
                f"Una vez emitido, envíe el recibo en formato *{EmailConfig.FORMATO_RECIBO}* como respuesta a este mismo correo, sin generar un nuevo hilo."
            )
            + self._crear_parrafo_html(
                f"Asimismo, se adjunta su contrato de locación de servicios correspondiente al periodo *{self._mes_inicio_contrato} - {self._mes_fin_contrato}*, firmado por el decano."
            )
        )

        return self._crear_encabezado_html() + cuerpo_contenido + self._crear_pie_html()


class EmailAdministrativoContratoBuilder(EmailAdministrativoBuilder):
    def _validar_datos(self) -> None:
        super()._validar_datos()
        if not self._mes_inicio_contrato or not self._mes_fin_contrato:
            raise ValueError("Periodo de contrato es requerido para primera vez")

    def construir_asunto(self) -> str:
        asunto_base = super().construir_asunto()
        return f"{asunto_base} | Contrato de locacion de servicios"

    def construir_cuerpo(self) -> str:
        self._validar_datos()

        texto_orden = (
            f"correspondiente al reconocimiento de deuda de *{self._mes} {self._anio}*"
            if self._es_reconocimiento_deuda
            else f"correspondiente al mes de *{self._mes} {self._anio}*.")

        if not texto_orden.endswith('.'):
            texto_orden += '.'

        cuerpo_contenido = (
            self._crear_parrafo_html(
                f"Adjunto su orden de servicio {texto_orden} Con este documento, ya puede proceder con la emisión de su recibo por honorarios. Para evitar retrasos en el pago, tenga en cuenta lo siguiente:"
            )
            + self._crear_lista_html([
                f"Número de RUC de entidad usuaria: *20148092282* (Número de RUC de la UNMSM).",
                "Pago: *al crédito.*",
                f"Plazo de vencimiento: *{EmailConfig.DIAS_VENCIMIENTO} días desde la emisión del recibo* en la plataforma de SUNAT.",
            ])
            + self._crear_parrafo_html(
                f"Una vez emitido, envíe el recibo en formato *{EmailConfig.FORMATO_RECIBO}* como respuesta a este mismo correo, sin generar un nuevo hilo."
            )
            + self._crear_parrafo_html(
                f"Asimismo, se adjunta su contrato de locación de servicios correspondiente al periodo *{self._mes_inicio_contrato} - {self._mes_fin_contrato}*, firmado por el decano."
            )
        )

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

    @staticmethod
    def crear_builder_contrato_primera_vez(tipo: TipoCorreo) -> EmailBuilder:
        if tipo == TipoCorreo.DOCENTE:
            return EmailDocenteContratoBuilder()
        elif tipo == TipoCorreo.ADMINISTRATIVO:
            return EmailAdministrativoContratoBuilder()
        else:
            raise ValueError(f"Tipo de correo no valido: {tipo}")
