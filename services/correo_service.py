from core.correos.envio_correos import (
    TipoCorreo,
    enviar_correo_contrato_primera_vez_desde_gui,
    enviar_lote_desde_gui_administrativos,
    enviar_lote_desde_gui_docentes,
    procesar_correo_individual_contrato_primera_vez,
    procesar_correos_administrativos_gmail,
    procesar_correos_docente_gmail,
)


def generar_data_correo_service(es_modo_contrato, tipo_var, ruta_excel, pdfs, pdf_orden):
    if es_modo_contrato:
        return [], procesar_correo_individual_contrato_primera_vez(
            ruta_excel=ruta_excel,
            hoja="list",
            pdf_orden_path=pdf_orden,
            tipo=TipoCorreo.DOCENTE if tipo_var == "Docente" else TipoCorreo.ADMINISTRATIVO,
        )

    if tipo_var == "Docente":
        return procesar_correos_docente_gmail(ruta_excel, "list", pdfs), None
    return procesar_correos_administrativos_gmail(ruta_excel, "list", pdfs), None


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
    if es_modo_contrato:
        return enviar_correo_contrato_primera_vez_desde_gui(
            datos_envio=data_envio_individual,
            pdf_contrato_path=pdf_contrato,
            mes=mes,
            mes_inicio_contrato=mes_inicio,
            mes_fin_contrato=mes_fin,
            tipo="docente" if tipo_var == "Docente" else "administrativo",
            anio=anio,
        )

    if tipo_var == "Docente":
        enviar_lote_desde_gui_docentes(data_envio, mes, anio, es_reconocimiento_deuda)
    else:
        enviar_lote_desde_gui_administrativos(data_envio, mes, anio, es_reconocimiento_deuda)

    return True
