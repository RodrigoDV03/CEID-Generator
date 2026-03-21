import os
import pickle
import logging
from typing import Optional

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.discovery import Resource

from .config import GmailConfig


# Configurar logging
logger = logging.getLogger(__name__)


class GmailAuthError(Exception):
    pass


class GmailService:
    
    def __init__(self, config: Optional[GmailConfig] = None):
        self.config = config or GmailConfig.default()
        self._service: Optional[Resource] = None
        self._credentials = None
    
    @property
    def service(self) -> Resource:
        if self._service is None:
            self._service = self._autenticar()
        return self._service
    
    def _cargar_credenciales(self) -> Optional[object]:
        if not os.path.exists(self.config.token_file):
            logger.info("No se encontró archivo de token existente")
            return None
        
        try:
            with open(self.config.token_file, 'rb') as token:
                creds = pickle.load(token)
                logger.info("Credenciales cargadas desde token.pickle")
                return creds
        except Exception as e:
            logger.warning(f"Error cargando token.pickle: {e}")
            return None
    
    def _guardar_credenciales(self, credentials: object) -> None:
        try:
            with open(self.config.token_file, 'wb') as token:
                pickle.dump(credentials, token)
            logger.info("Credenciales guardadas en token.pickle")
        except Exception as e:
            logger.error(f"Error guardando credenciales: {e}")
            raise GmailAuthError(f"No se pudieron guardar las credenciales: {e}")
    
    def _refrescar_credenciales(self, credentials: object) -> bool:
        try:
            if credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
                logger.info("Credenciales refrescadas exitosamente")
                return True
            return False
        except Exception as e:
            logger.warning(f"Error refrescando credenciales: {e}")
            return False
    
    def _autenticacion_oauth(self) -> object:
        if not os.path.exists(self.config.credentials_file):
            raise GmailAuthError(
                f"No se encontró {self.config.credentials_file}. "
                "Descarga las credenciales OAuth2 desde Google Cloud Console."
            )
        
        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                self.config.credentials_file,
                self.config.scopes
            )
            creds = flow.run_local_server(port=0)
            logger.info("Autenticación OAuth2 completada exitosamente")
            return creds
        except Exception as e:
            logger.error(f"Error en flujo OAuth2: {e}")
            raise GmailAuthError(f"Falló la autenticación OAuth2: {e}")
    
    def _autenticar(self) -> Resource:
        """
        Autentica y construye el servicio de Gmail.
        
        Maneja el flujo completo:
        1. Cargar credenciales guardadas
        2. Validar si son válidas
        3. Refrescar si están expiradas
        4. Autenticar si no hay credenciales válidas
        5. Guardar credenciales nuevas/actualizadas
        """
        creds = self._cargar_credenciales()
        
        # Si no hay credenciales o no son válidas
        if not creds or not creds.valid:
            if creds:
                # Intentar refrescar
                if not self._refrescar_credenciales(creds):
                    # Si falla el refresco, autenticar de nuevo
                    creds = self._autenticacion_oauth()
            else:
                # No hay credenciales, autenticar
                creds = self._autenticacion_oauth()
            
            # Guardar las credenciales nuevas/actualizadas
            self._guardar_credenciales(creds)
        
        # Construir el servicio
        try:
            service = build('gmail', 'v1', credentials=creds)
            self._credentials = creds
            logger.info("Servicio de Gmail construido exitosamente")
            return service
        except Exception as e:
            logger.error(f"Error construyendo servicio de Gmail: {e}")
            raise GmailAuthError(f"No se pudo construir el servicio de Gmail: {e}")
    
    def obtener_firma(self) -> str:
        """
        Obtiene la firma de Gmail configurada en la cuenta.
        
        Busca la firma principal de la cuenta de Gmail.
        """
        try:
            send_as = self.service.users().settings().sendAs().list(userId='me').execute()
            
            for alias in send_as.get('sendAs', []):
                if alias.get('isPrimary'):
                    signature = alias.get('signature', '')
                    if signature:
                        logger.info("Firma de Gmail obtenida exitosamente")
                        return signature
            
            logger.info("No se encontró firma de Gmail configurada")
            return ""
            
        except Exception as e:
            logger.warning(f"No se pudo obtener la firma de Gmail: {e}")
            return ""
    
    def es_autenticado(self) -> bool:
        """
        Verifica si el servicio está autenticado y listo para usar.
        
        Returns:
            bool: True si está autenticado, False en caso contrario
        """
        return self._service is not None and self._credentials is not None
    
    def cerrar_sesion(self) -> None:
        """
        Cierra la sesión y limpia las credenciales guardadas.
        
        Elimina el archivo token para forzar nueva autenticación en el próximo uso.
        """
        self._service = None
        self._credentials = None
        
        if os.path.exists(self.config.token_file):
            try:
                os.remove(self.config.token_file)
                logger.info("Sesión cerrada y token eliminado")
            except Exception as e:
                logger.error(f"Error eliminando token: {e}")


# Función de compatibilidad con código legacy
def autenticar_gmail() -> Resource:
    logger.debug("autenticar_gmail() está deprecated. Usar GmailService en su lugar.")
    service_manager = GmailService()
    return service_manager.service


def obtener_firma_gmail(service: Resource) -> str:
    logger.debug("obtener_firma_gmail() está deprecated. Usar GmailService.obtener_firma() en su lugar.")
    try:
        send_as = service.users().settings().sendAs().list(userId='me').execute()
        
        for alias in send_as.get('sendAs', []):
            if alias.get('isPrimary'):
                signature = alias.get('signature', '')
                if signature:
                    return signature
        return ""
        
    except Exception as e:
        logger.warning(f"No se pudo obtener la firma de Gmail: {e}")
        return ""
