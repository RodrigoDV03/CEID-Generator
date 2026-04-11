# ===== IMPORTS DE MÓDULOS ESPECIALIZADOS =====

from .config import (
    AÑO_ACTUAL,
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

# ===== EXPORTS =====

__all__ = [
    # Constantes y configuración
    'AÑO_ACTUAL',
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