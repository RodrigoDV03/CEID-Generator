from core.fases.fase_final.generador_fase_final import procesar_planilla_fase_final
from core.fases.fase_inicial.generador_fase_inicial import procesar_planilla_fase_inicial


def procesar_fase_inicial_service(planilla_path, hoja, carpeta_destino, mes, numero_armada, tipo_fase):
    return procesar_planilla_fase_inicial(planilla_path, hoja, carpeta_destino, mes, numero_armada, tipo_fase)


def procesar_fase_final_service(planilla_path, excel_control_pagos, hoja, carpeta_salida, mes, numero_armada, tipo_fase):
    return procesar_planilla_fase_final(
        planilla_path,
        excel_control_pagos,
        hoja,
        carpeta_salida,
        mes,
        numero_armada,
        tipo_fase,
    )
