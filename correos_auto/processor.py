import os
import logging
from typing import List, Dict, Optional
from dataclasses import dataclass
import pandas as pd

from .config import TipoCorreo
from .pdf_extractor import PDFExtractor
from .gmail_service import GmailService


logger = logging.getLogger(__name__)


class ExcelValidationError(Exception):
    """Excepción para errores de validación de Excel."""
    pass


@dataclass
class DatosEnvio:
    nombre: str
    correo: str
    pdf_path: str
    servicio: Optional[str] = None
    
    def to_dict(self) -> Dict:
        """Convierte a diccionario para compatibilidad."""
        return {
            "nombre": self.nombre,
            "correo": self.correo,
            "pdf_path": self.pdf_path,
            "servicio": self.servicio
        }


class ExcelReader:
    """
    Lector y validador de archivos Excel.
    
    Maneja la lectura de archivos Excel y validación de columnas requeridas.
    """
    
    COLUMNAS_REQUERIDAS = ['Docente', 'Correo Institucional', 'N_Ruc']
    
    def __init__(self, ruta_excel: str, hoja: str):
        if not os.path.exists(ruta_excel):
            raise FileNotFoundError(f"No se encontró el archivo Excel: {ruta_excel}")
        
        self.ruta_excel = ruta_excel
        self.hoja = hoja
        self._df: Optional[pd.DataFrame] = None
    
    @property
    def df(self) -> pd.DataFrame:
        if self._df is None:
            self._df = self._leer_y_validar()
        return self._df
    
    def _leer_y_validar(self) -> pd.DataFrame:
        try:
            df = pd.read_excel(self.ruta_excel, sheet_name=self.hoja)
            df.columns = df.columns.str.strip()
            
            # Validar columnas requeridas
            columnas_faltantes = set(self.COLUMNAS_REQUERIDAS) - set(df.columns)
            if columnas_faltantes:
                raise ExcelValidationError(
                    f"Faltan las siguientes columnas en el Excel: {columnas_faltantes}"
                )
            
            logger.info(f"Excel leído exitosamente: {len(df)} registros")
            return df
            
        except Exception as e:
            logger.error(f"Error leyendo Excel: {e}")
            raise ExcelValidationError(f"Error leyendo el archivo Excel: {e}")
    
    def buscar_por_nombre(self, nombre: str) -> Optional[pd.Series]:
        """Busca una fila por nombre exacto."""
        fila = self.df[self.df['Docente'] == nombre]
        if not fila.empty:
            return fila.iloc[0]
        return None
    
    def buscar_por_ruc(self, ruc: str) -> Optional[pd.Series]:
        """Busca una fila por RUC."""
        # Convertir RUC a string y eliminar espacios
        ruc_limpio = str(ruc).strip()
        
        # Normalizar RUCs del DataFrame
        def normalizar_ruc(valor):
            """Convierte el valor a string limpio sin decimales."""
            if pd.isna(valor):
                return ''
            try:
                # Si es numérico, convertir a int para eliminar decimales
                return str(int(float(valor)))
            except (ValueError, TypeError):
                # Si no es numérico, devolver como string limpio
                return str(valor).strip()
        
        df_rucs_normalizados = self.df['N_Ruc'].apply(normalizar_ruc)
        
        # Buscar coincidencia exacta
        fila = self.df[df_rucs_normalizados == ruc_limpio]
        
        if not fila.empty:
            return fila.iloc[0]
        return None


class PDFProcessor:
    """Procesador de PDFs para extraer RUC y servicios y emparejarlos con Excel."""
    
    def __init__(
        self,
        excel_reader: ExcelReader,
        incluir_servicio: bool = True,
        debug: bool = False
    ):
        self.excel_reader = excel_reader
        self.incluir_servicio = incluir_servicio
        self.debug = debug
    
    def procesar_pdf(self, pdf_path: str) -> Optional[DatosEnvio]:
        # Extraer RUC del PDF (identificador único)
        extractor = PDFExtractor(pdf_path)
        ruc_extraido = extractor.extraer_ruc(debug=self.debug)
        
        if not ruc_extraido:
            logger.warning(f"No se encontró RUC en {os.path.basename(pdf_path)}")
            print(f"⚠ No se encontró RUC en {os.path.basename(pdf_path)}, omitido.")
            return None
        
        # Buscar en Excel por RUC
        fila = self.excel_reader.buscar_por_ruc(ruc_extraido)
        
        if fila is None:
            logger.warning(f"No se encontró coincidencia para RUC {ruc_extraido} ({os.path.basename(pdf_path)})")
            print(f"⚠ No se encontró coincidencia para RUC {ruc_extraido} (archivo: {os.path.basename(pdf_path)}), omitido.")
            return None
        
        # Obtener datos del Excel
        nombre_excel = fila['Docente']
        correo = fila['Correo Institucional']
        
        # Extraer servicio si es necesario
        servicio = None
        if self.incluir_servicio:
            servicio = extractor.extraer_servicios(debug=self.debug)
            if not servicio:
                logger.warning(f"No se encontró servicio para {nombre_excel}")
                print(f"⚠ No se encontró servicio para {nombre_excel}.")
                return None
        
        # Crear datos de envío
        datos = DatosEnvio(
            nombre=nombre_excel,
            correo=correo,
            pdf_path=pdf_path,
            servicio=servicio
        )
        
        # Log de información
        servicio_info = f" - {servicio}" if servicio else ""
        print(f"✓ RUC {ruc_extraido}: {nombre_excel} - {correo}{servicio_info}")
        
        return datos
    
    def procesar_lote(self, lista_pdfs: List[str]) -> List[DatosEnvio]:
        resultados = []
        
        logger.info(f"Procesando {len(lista_pdfs)} PDFs...")
        
        for pdf_path in lista_pdfs:
            datos = self.procesar_pdf(pdf_path)
            if datos:
                resultados.append(datos)
        
        logger.info(f"Procesamiento completado: {len(resultados)}/{len(lista_pdfs)} PDFs procesados")
        return resultados


class CorreosProcessor:
    
    def __init__(self, gmail_service: Optional[GmailService] = None):
        self.gmail_service = gmail_service or GmailService()
    
    def procesar_correos(
        self,
        ruta_excel: str,
        hoja: str,
        lista_pdfs: List[str],
        tipo: TipoCorreo,
        debug: bool = False
    ) -> List[Dict]:
        # Autenticar Gmail
        try:
            _ = self.gmail_service.service  # Forzar autenticación
        except Exception as e:
            logger.error(f"Error autenticando con Gmail: {e}")
            print(f"❌ Error autenticando con Gmail API: {e}")
            return []
        
        # Leer Excel
        excel_reader = ExcelReader(ruta_excel, hoja)
        
        # Procesar PDFs
        incluir_servicio = (tipo == TipoCorreo.DOCENTE)
        processor = PDFProcessor(excel_reader, incluir_servicio, debug)
        resultados = processor.procesar_lote(lista_pdfs)
        
        # Convertir a diccionarios para compatibilidad
        return [datos.to_dict() for datos in resultados]


# Funciones de compatibilidad con código legacy
def procesar_correos_docente_gmail(
    ruta_excel: str,
    hoja: str,
    lista_pdfs: List[str]
) -> List[Dict]:
    logger.debug("procesar_correos_docente_gmail() está deprecated. Usar CorreosProcessor.")
    
    processor = CorreosProcessor()
    return processor.procesar_correos(ruta_excel, hoja, lista_pdfs, TipoCorreo.DOCENTE)


def procesar_correos_administrativos_gmail(
    ruta_excel: str,
    hoja: str,
    lista_pdfs: List[str],
    debug: bool = False
) -> List[Dict]:
    logger.debug("procesar_correos_administrativos_gmail() está deprecated. Usar CorreosProcessor.")
    
    processor = CorreosProcessor()
    return processor.procesar_correos(ruta_excel, hoja, lista_pdfs, TipoCorreo.ADMINISTRATIVO, debug)


def enviar_lote_docentes_gmail(ruta_excel: str, hoja: str, lista_pdfs: List[str], mes: str) -> None:
    logger.debug("enviar_lote_docentes_gmail() está deprecated.")
    
    from .email_sender import enviar_lote_desde_gui_docentes
    
    resultados = procesar_correos_docente_gmail(ruta_excel, hoja, lista_pdfs)
    enviar_lote_desde_gui_docentes(resultados, mes)


def enviar_lote_administrativos_gmail(ruta_excel: str, hoja: str, lista_pdfs: List[str], mes: str) -> None:
    logger.debug("enviar_lote_administrativos_gmail() está deprecated.")
    
    from .email_sender import enviar_lote_desde_gui_administrativos
    
    resultados = procesar_correos_administrativos_gmail(ruta_excel, hoja, lista_pdfs)
    enviar_lote_desde_gui_administrativos(resultados, mes)
