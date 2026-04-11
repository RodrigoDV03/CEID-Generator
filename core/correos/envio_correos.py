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
)

from .email_builder import (
    EmailBuilder,
    EmailDocenteBuilder,
    EmailAdministrativoBuilder,
    EmailDocenteContratoBuilder,
    EmailAdministrativoContratoBuilder,
    EmailBuilderFactory,
)

from .pdf_extractor import (
    PDFExtractor,
    ServicioExtractor,
    RUCExtractor,
    tiene_contrato_locacion
)

from .email_sender import (
    GmailEmailSender,
    EmailPersonalizado,
    LoteEmailSender,
    EmailSendError,
)

from .processor import (
    CorreosProcessor,
    PDFProcessor,
    ExcelReader,
    DatosEnvio,
    ExcelValidationError,
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
    
    # Email Builder
    'EmailBuilder',
    'EmailDocenteBuilder',
    'EmailAdministrativoBuilder',
    'EmailDocenteContratoBuilder',
    'EmailAdministrativoContratoBuilder',
    'EmailBuilderFactory',
    
    # PDF Extractor
    'PDFExtractor',
    'ServicioExtractor',
    'RUCExtractor',
    'tiene_contrato_locacion',
    
    # Email Sender
    'GmailEmailSender',
    'EmailPersonalizado',
    'LoteEmailSender',
    'EmailSendError',
    
    # Processor
    'CorreosProcessor',
    'PDFProcessor',
    'ExcelReader',
    'DatosEnvio',
    'ExcelValidationError',
]

if __name__ == '__main__':
    # Ejemplo de uso
    print("=" * 60)
    print("Sistema de Envío Automático de Correos")
    print("=" * 60)