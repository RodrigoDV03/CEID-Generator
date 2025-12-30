import os
import base64
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional

from googleapiclient.discovery import Resource
from googleapiclient.errors import HttpError

from .config import TipoCorreo, AÑO_ACTUAL
from .gmail_service import GmailService
from .email_builder import EmailBuilderFactory
from .gmail_service import autenticar_gmail


logger = logging.getLogger(__name__)


class EmailSendError(Exception):
    pass


class GmailMessageBuilder:
    def __init__(self, destinatario: str, asunto: str):
        self.destinatario = destinatario
        self.asunto = asunto
        self.mensaje = MIMEMultipart()
        self.mensaje['To'] = destinatario
        self.mensaje['Subject'] = asunto
    
    def agregar_cuerpo_html(self, cuerpo_html: str) -> 'GmailMessageBuilder':
        self.mensaje.attach(MIMEText(cuerpo_html, 'html'))
        return self
    
    def agregar_adjunto_pdf(self, pdf_path: str) -> 'GmailMessageBuilder':
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"No se encontró el archivo PDF: {pdf_path}")
        
        try:
            with open(pdf_path, 'rb') as adjunto:
                parte = MIMEBase('application', 'octet-stream')
                parte.set_payload(adjunto.read())
                encoders.encode_base64(parte)
                parte.add_header(
                    'Content-Disposition',
                    f'attachment; filename={os.path.basename(pdf_path)}'
                )
                self.mensaje.attach(parte)
        except Exception as e:
            logger.error(f"Error adjuntando PDF {pdf_path}: {e}")
            raise EmailSendError(f"No se pudo adjuntar el archivo PDF: {e}")
        
        return self
    
    def codificar(self) -> str:
        return base64.urlsafe_b64encode(self.mensaje.as_bytes()).decode()


class GmailEmailSender:
    def __init__(self, gmail_service: GmailService):
        self.gmail_service = gmail_service
    
    def enviar_con_firma(
        self,
        destinatario: str,
        asunto: str,
        cuerpo_html: str,
        pdf_path: str
    ) -> dict:
        try:
            # Crear mensaje MIME
            builder = GmailMessageBuilder(destinatario, asunto)
            raw_message = (builder
                          .agregar_cuerpo_html(cuerpo_html)
                          .agregar_adjunto_pdf(pdf_path)
                          .codificar())
            
            # Crear draft (esto permite que Gmail añada la firma)
            draft_body = {'message': {'raw': raw_message}}
            draft = self.gmail_service.service.users().drafts().create(
                userId='me',
                body=draft_body
            ).execute()
            
            logger.info(f"Draft creado con ID: {draft['id']}")
            
            # Enviar el draft
            result = self.gmail_service.service.users().drafts().send(
                userId='me',
                body={'id': draft['id']}
            ).execute()
            
            logger.info(f"Correo enviado exitosamente a {destinatario}")
            return result
            
        except HttpError as e:
            logger.error(f"Error HTTP enviando correo a {destinatario}: {e}")
            raise EmailSendError(f"Error de Gmail API: {e}")
        except Exception as e:
            logger.error(f"Error inesperado enviando correo a {destinatario}: {e}")
            raise EmailSendError(f"Error inesperado: {e}")
    
    def enviar(
        self,
        destinatario: str,
        asunto: str,
        cuerpo_html: str,
        pdf_path: str,
        nombre_destinatario: str
    ) -> bool:
        try:
            self.enviar_con_firma(destinatario, asunto, cuerpo_html, pdf_path)
            print(f"Correo enviado a {nombre_destinatario}")
            return True
        except EmailSendError as e:
            print(f"❌ Error enviando correo a {nombre_destinatario}: {e}")
            return False
        except Exception as e:
            print(f"❌ Error inesperado enviando correo a {nombre_destinatario}: {e}")
            return False


class EmailPersonalizado:
    def __init__(self, gmail_service: GmailService):
        """
        Inicializa el enviador de correos personalizados.
        
        Args:
            gmail_service: Instancia de GmailService autenticado
        """
        self.gmail_service = gmail_service
        self.sender = GmailEmailSender(gmail_service)
        self._firma_html: Optional[str] = None
    
    @property
    def firma_html(self) -> str:
        """
        Obtiene la firma HTML de Gmail (con caching).
        
        Returns:
            str: Firma HTML o cadena vacía
        """
        if self._firma_html is None:
            self._firma_html = self.gmail_service.obtener_firma()
        return self._firma_html
    
    def enviar(
        self,
        nombre: str,
        pdf_path: str,
        destinatario: str,
        mes: str,
        tipo: TipoCorreo,
        servicio: Optional[str] = None,
        anio: Optional[int] = None
    ) -> bool:
        if anio is None:
            anio = AÑO_ACTUAL
        
        # Validar que docentes tengan servicio
        if tipo == TipoCorreo.DOCENTE and not servicio:
            print(f"⚠ No se puede enviar correo a {nombre}: servicio no especificado")
            return False
        
        try:
            # Construir el correo según el tipo
            builder = EmailBuilderFactory.crear_builder(tipo)
            builder.con_mes(mes).con_anio(anio).con_firma(self.firma_html).con_nombre(nombre)
            
            if tipo == TipoCorreo.DOCENTE:
                builder.con_servicio(servicio)
            
            asunto = builder.construir_asunto()
            cuerpo_html = builder.construir_cuerpo()
            
            # Enviar el correo
            return self.sender.enviar(destinatario, asunto, cuerpo_html, pdf_path, nombre)
            
        except Exception as e:
            print(f"❌ Error preparando correo para {nombre}: {e}")
            logger.error(f"Error en enviar correo personalizado para {nombre}: {e}")
            return False


class LoteEmailSender:
    def __init__(self, gmail_service: GmailService):
        self.email_personalizado = EmailPersonalizado(gmail_service)
    
    def enviar_lote(
        self,
        data_para_envio: list,
        mes: str,
        tipo: TipoCorreo,
        anio: Optional[int] = None
    ) -> dict:
        if anio is None:
            anio = AÑO_ACTUAL
        
        tipo_texto = "docentes" if tipo == TipoCorreo.DOCENTE else "personal administrativo"
        print(f"\nEnviando {len(data_para_envio)} correos a {tipo_texto}...")
        
        exitosos = 0
        fallidos = 0
        
        for datos in data_para_envio:
            servicio = datos.get("servicio") if tipo == TipoCorreo.DOCENTE else None
            
            resultado = self.email_personalizado.enviar(
                nombre=datos["nombre"],
                pdf_path=datos["pdf_path"],
                destinatario=datos["correo"],
                mes=mes,
                tipo=tipo,
                servicio=servicio,
                anio=anio
            )
            
            if resultado:
                exitosos += 1
            else:
                fallidos += 1
        
        print(f"\n✅ Envío completado: {exitosos} exitosos, {fallidos} fallidos")
        
        return {
            "exitosos": exitosos,
            "fallidos": fallidos,
            "total": len(data_para_envio)
        }


# Funciones de compatibilidad con código legacy
def crear_mensaje_gmail_con_firma(
    service: Resource,
    destinatario: str,
    asunto: str,
    cuerpo_html: str,
    pdf_path: str
) -> dict:
    logger.debug("crear_mensaje_gmail_con_firma() está deprecated. Usar GmailEmailSender.")
    
    # Crear un GmailService temporal con el servicio existente
    gmail_service = GmailService()
    gmail_service._service = service
    
    sender = GmailEmailSender(gmail_service)
    return sender.enviar_con_firma(destinatario, asunto, cuerpo_html, pdf_path)


def enviar_correo_gmail(
    service: Resource,
    destinatario: str,
    asunto: str,
    cuerpo_html: str,
    pdf_path: str,
    nombre: str
) -> None:
    logger.warning("enviar_correo_gmail() está deprecated. Usar GmailEmailSender.")
    
    gmail_service = GmailService()
    gmail_service._service = service
    
    sender = GmailEmailSender(gmail_service)
    sender.enviar(destinatario, asunto, cuerpo_html, pdf_path, nombre)


def enviar_correo_personalizado(
    service: Resource,
    nombre: str,
    pdf_path: str,
    destinatario: str,
    mes: str,
    tipo: str,
    servicio: Optional[str] = None
) -> None:
    logger.debug("enviar_correo_personalizado() está deprecated. Usar EmailPersonalizado.")
    
    gmail_service = GmailService()
    gmail_service._service = service
    
    # Convertir string a TipoCorreo enum
    tipo_enum = TipoCorreo.DOCENTE if tipo == "docente" else TipoCorreo.ADMINISTRATIVO
    
    email_personalizado = EmailPersonalizado(gmail_service)
    email_personalizado.enviar(nombre, pdf_path, destinatario, mes, tipo_enum, servicio)


def enviar_lote_desde_gui(data_para_envio: list, mes: str, tipo: str, anio: Optional[int] = None) -> None:
    logger.debug("enviar_lote_desde_gui() está deprecated. Usar LoteEmailSender.")
    
    service = autenticar_gmail()
    
    gmail_service = GmailService()
    gmail_service._service = service
    
    # Convertir string a TipoCorreo enum
    tipo_enum = TipoCorreo.DOCENTE if tipo == "docente" else TipoCorreo.ADMINISTRATIVO
    
    lote_sender = LoteEmailSender(gmail_service)
    lote_sender.enviar_lote(data_para_envio, mes, tipo_enum)


def enviar_lote_desde_gui_docentes(data_para_envio: list, mes: str, anio: Optional[int] = None) -> None:
    """Wrapper para mantener compatibilidad con código existente."""
    enviar_lote_desde_gui(data_para_envio, mes, "docente", anio)


def enviar_lote_desde_gui_administrativos(data_para_envio: list, mes: str, anio: Optional[int] = None) -> None:
    """Wrapper para mantener compatibilidad con código existente."""
    enviar_lote_desde_gui(data_para_envio, mes, "administrativo", anio)
