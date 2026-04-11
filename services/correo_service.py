from core.correos.config import TipoCorreo
from core.correos.email_sender import EmailPersonalizado, LoteEmailSender
from core.correos.gmail_service import GmailService
from core.correos.processor import CorreosProcessor, ExcelReader, PDFProcessor


def generar_data_correo_service(es_modo_contrato, tipo_var, ruta_excel, pdfs, pdf_orden):
    gmail_service = GmailService()
    processor = CorreosProcessor(gmail_service)

    if es_modo_contrato:
        excel_reader = ExcelReader(ruta_excel, "list")
        pdf_processor = PDFProcessor(
            excel_reader,
            incluir_servicio=tipo_var == "Docente",
            incluir_modalidad=tipo_var == "Docente",
        )
        datos = pdf_processor.procesar_orden_individual_contrato_primera_vez(pdf_orden)
        return [], datos.to_dict() if datos else None

    tipo_correo = TipoCorreo.DOCENTE if tipo_var == "Docente" else TipoCorreo.ADMINISTRATIVO
    return processor.procesar_correos(ruta_excel, "list", pdfs, tipo_correo), None


def enviar_correos_service(
    es_modo_contrato,
    tipo_var,
    data_envio,
    data_envio_individual,
    pdf_contrato,
    mes,
    mes_inicio,
    mes_fin,
    anio,
    es_reconocimiento_deuda=False,
):
    gmail_service = GmailService()
    lote_sender = LoteEmailSender(gmail_service)
    correo_individual = EmailPersonalizado(gmail_service)

    if es_modo_contrato:
        return correo_individual.enviar_contrato_primera_vez(
            nombre=data_envio_individual["nombre"],
            pdf_orden_path=data_envio_individual["pdf_path"],
            pdf_contrato_path=pdf_contrato,
            destinatario=data_envio_individual["correo"],
            mes=mes,
            tipo=TipoCorreo.DOCENTE if tipo_var == "Docente" else TipoCorreo.ADMINISTRATIVO,
            mes_inicio_contrato=mes_inicio,
            mes_fin_contrato=mes_fin,
            servicio=data_envio_individual.get("servicio"),
            modalidad=data_envio_individual.get("modalidad"),
            anio=anio,
        )

    lote_sender.enviar_lote(
        data_envio,
        mes,
        TipoCorreo.DOCENTE if tipo_var == "Docente" else TipoCorreo.ADMINISTRATIVO,
        anio,
        es_reconocimiento_deuda,
    )

    return True
