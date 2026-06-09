from core.correos.config import TipoCorreo
from core.correos.email_builder import EmailBuilderFactory
from core.correos.email_sender import EmailPersonalizado, GmailEmailSender, LoteEmailSender
from core.correos.gmail_service import GmailService
from core.correos.processor import CorreosProcessor, ExcelReader, PDFProcessor
from typing import Any


_gmail_service_shared = None


def _resolver_tipo_correo(tipo_var: str) -> TipoCorreo:
    return TipoCorreo.DOCENTE if tipo_var == "Docente" else TipoCorreo.ADMINISTRATIVO


def _obtener_gmail_service() -> GmailService:
    global _gmail_service_shared

    if _gmail_service_shared is None:
        _gmail_service_shared = GmailService()

    return _gmail_service_shared


def generar_data_correo_service(es_modo_contrato, tipo_var, ruta_excel, pdfs, pdf_orden):
    gmail_service = _obtener_gmail_service()
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

    tipo_correo = _resolver_tipo_correo(tipo_var)
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
    gmail_service = _obtener_gmail_service()
    lote_sender = LoteEmailSender(gmail_service)
    correo_individual = EmailPersonalizado(gmail_service)

    if es_modo_contrato:
        resultado = correo_individual.enviar_contrato_primera_vez(
            nombre=data_envio_individual["nombre"],
            pdf_orden_path=data_envio_individual["pdf_path"],
            pdf_contrato_path=pdf_contrato,
            destinatario=data_envio_individual["correo"],
            mes=mes,
            tipo=_resolver_tipo_correo(tipo_var),
            mes_inicio_contrato=mes_inicio,
            mes_fin_contrato=mes_fin,
            servicio=data_envio_individual.get("servicio"),
            modalidad=data_envio_individual.get("modalidad"),
            anio=anio,
        )

        return {
            "exitosos": 1 if resultado.success else 0,
            "fallidos": 0 if resultado.success else 1,
            "total": 1,
            "thread_ids": [resultado.thread_id] if resultado.success and resultado.thread_id else [],
            "resultados": [
                {
                    "nombre": data_envio_individual["nombre"],
                    "destinatario": data_envio_individual["correo"],
                    "success": resultado.success,
                    "message_id": resultado.message_id,
                    "thread_id": resultado.thread_id,
                    "error": resultado.error,
                }
            ],
        }

    resumen = lote_sender.enviar_lote(
        data_envio,
        mes,
        _resolver_tipo_correo(tipo_var),
        anio,
        es_reconocimiento_deuda,
    )

    return resumen


def previsualizar_correos_service(
    data_envio,
    mes,
    tipo_var,
    anio,
    es_reconocimiento_deuda=False,
    es_modo_contrato=False,
    mes_inicio=None,
    mes_fin=None,
    pdf_contrato=None,
):
    """
    Genera asunto y cuerpo HTML para cada correo sin enviarlo.

    Retorna una lista de diccionarios con:
    asunto, cuerpo_html, pdf_path, pdf_paths, destinatario, nombre.
    """
    tipo = _resolver_tipo_correo(tipo_var)
    previsualizaciones = []
    for datos in data_envio:
        if es_modo_contrato:
            builder: Any = EmailBuilderFactory.crear_builder_contrato_primera_vez(tipo)
            builder.con_periodo_contrato(mes_inicio or "", mes_fin or "")
        else:
            builder: Any = EmailBuilderFactory.crear_builder(tipo)
            builder.con_reconocimiento_deuda(es_reconocimiento_deuda)

        # La firma se deja a Gmail en el momento del envio.
        builder.con_mes(mes).con_anio(anio).con_firma("").con_nombre(datos["nombre"])

        if tipo == TipoCorreo.DOCENTE:
            builder.con_servicio(datos.get("servicio", ""))
            builder.con_modalidad(datos.get("modalidad", ""))

        pdf_paths = [datos["pdf_path"]]
        if es_modo_contrato and pdf_contrato:
            pdf_paths.append(pdf_contrato)

        previsualizaciones.append(
            {
                "asunto": builder.construir_asunto(),
                "cuerpo_html": builder.construir_cuerpo(),
                "pdf_path": datos["pdf_path"],
                "pdf_paths": pdf_paths,
                "destinatario": datos["correo"],
                "nombre": datos["nombre"],
            }
        )

    return previsualizaciones


def enviar_previsualizaciones_service(previsualizaciones_editadas):
    """
    Envia correos usando asunto/cuerpo ya editados en previsualizacion.

    Cada item debe incluir: destinatario, asunto, cuerpo_html y pdf_paths.
    """
    gmail_service = _obtener_gmail_service()
    sender = GmailEmailSender(gmail_service)

    exitosos = 0
    fallidos = 0
    resultados = []

    for item in previsualizaciones_editadas:
        try:
            resultado = sender.enviar_con_firma(
                destinatario=item["destinatario"],
                asunto=item["asunto"],
                cuerpo_html=item["cuerpo_html"],
                pdf_paths=item.get("pdf_paths") or [item.get("pdf_path")],
            )
            if resultado.success:
                print(f"Correo enviado a {item.get('nombre', item['destinatario'])}")
                exitosos += 1
            else:
                fallidos += 1

            resultados.append(
                {
                    "nombre": item.get("nombre"),
                    "destinatario": item["destinatario"],
                    "success": resultado.success,
                    "message_id": resultado.message_id,
                    "thread_id": resultado.thread_id,
                    "error": resultado.error,
                }
            )
        except Exception as e:
            print(f"Error enviando correo a {item.get('nombre', item['destinatario'])}: {e}")
            fallidos += 1
            resultados.append(
                {
                    "nombre": item.get("nombre"),
                    "destinatario": item["destinatario"],
                    "success": False,
                    "message_id": None,
                    "thread_id": None,
                    "error": str(e),
                }
            )

    resultados = locals().get("resultados", [])

    return {
        "exitosos": exitosos,
        "fallidos": fallidos,
        "total": len(previsualizaciones_editadas),
        "thread_ids": [r["thread_id"] for r in resultados if r["success"] and r["thread_id"]],
        "resultados": resultados,
    }
