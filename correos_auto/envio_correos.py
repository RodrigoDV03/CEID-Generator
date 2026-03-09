# ===== IMPORTS DE MÓDULOS ESPECIALIZADOS =====

from .config import (
    AÑO_ACTUAL,
    GmailConfig,
    TipoCorreo,
    PatronesRegex,
    ServiciosConfig,
    EmailConfig,
    get_app_dir
)

from .gmail_service import (
    GmailService,
    GmailAuthError,
    autenticar_gmail,  # Legacy
    obtener_firma_gmail  # Legacy
)

from .email_builder import (
    EmailBuilder,
    EmailDocenteBuilder,
    EmailAdministrativoBuilder,
    EmailBuilderFactory,
    generar_cuerpo_correo_docente_html,  # Legacy
    generar_cuerpo_correo_administrativo_html  # Legacy
)

from .pdf_extractor import (
    PDFExtractor,
    ServicioExtractor,
    RUCExtractor,
    extraer_ruc,  # Legacy
    extraer_servicios  # Legacy
)

from .email_sender import (
    GmailEmailSender,
    EmailPersonalizado,
    LoteEmailSender,
    EmailSendError,
    crear_mensaje_gmail_con_firma,  # Legacy
    enviar_correo_gmail,  # Legacy
    enviar_correo_personalizado,  # Legacy
    enviar_lote_desde_gui,  # Legacy
    enviar_lote_desde_gui_docentes,  # Legacy
    enviar_lote_desde_gui_administrativos  # Legacy
)

from .processor import (
    CorreosProcessor,
    PDFProcessor,
    ExcelReader,
    DatosEnvio,
    ExcelValidationError,
    procesar_correos_docente_gmail,  # Legacy
    procesar_correos_administrativos_gmail,  # Legacy
    enviar_lote_docentes_gmail,  # Legacy
    enviar_lote_administrativos_gmail  # Legacy
)


# ===== VARIABLES GLOBALES (para compatibilidad) =====

año_actual = AÑO_ACTUAL

# Configuración de Gmail (para compatibilidad)
GMAIL_CONFIG = {
    "scopes": GmailConfig.default().scopes,
    "credentials_file": GmailConfig.default().credentials_file,
    "token_file": GmailConfig.default().token_file,
    "remitente": GmailConfig.default().remitente
}


# ===== EXPORTS =====

__all__ = [
    # Constantes y configuración
    'AÑO_ACTUAL',
    'año_actual',
    'GMAIL_CONFIG',
    'GmailConfig',
    'TipoCorreo',
    'PatronesRegex',
    'ServiciosConfig',
    'EmailConfig',
    'get_app_dir',
    
    # Gmail Service
    'GmailService',
    'GmailAuthError',
    'autenticar_gmail',
    'obtener_firma_gmail',
    
    # Email Builder
    'EmailBuilder',
    'EmailDocenteBuilder',
    'EmailAdministrativoBuilder',
    'EmailBuilderFactory',
    'generar_cuerpo_correo_docente_html',
    'generar_cuerpo_correo_administrativo_html',
    
    # PDF Extractor
    'PDFExtractor',
    'ServicioExtractor',
    'RUCExtractor',
    'extraer_ruc',
    'extraer_servicios',
    
    # Email Sender
    'GmailEmailSender',
    'EmailPersonalizado',
    'LoteEmailSender',
    'EmailSendError',
    'crear_mensaje_gmail_con_firma',
    'enviar_correo_gmail',
    'enviar_correo_personalizado',
    'enviar_lote_desde_gui',
    'enviar_lote_desde_gui_docentes',
    'enviar_lote_desde_gui_administrativos',
    
    # Processor
    'CorreosProcessor',
    'PDFProcessor',
    'ExcelReader',
    'DatosEnvio',
    'ExcelValidationError',
    'procesar_correos_docente_gmail',
    'procesar_correos_administrativos_gmail',
    'enviar_lote_docentes_gmail',
    'enviar_lote_administrativos_gmail',
]

if __name__ == '__main__':
    # Ejemplo de uso
    print("=" * 60)
    print("Sistema de Envío Automático de Correos")
    print("=" * 60)