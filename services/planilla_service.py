from core.planillas.generador_planilla import generar_planilla


def generar_planilla_service(
    archivo_cursos,
    archivo_docentes,
    archivo_clasif,
    archivo_coordinacion,
    mes,
    bono,
    carpeta_destino,
):
    return generar_planilla(
        archivo_cursos,
        archivo_docentes,
        archivo_clasif,
        archivo_coordinacion,
        mes,
        bono,
        carpeta_destino,
    )
