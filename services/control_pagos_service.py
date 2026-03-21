from core.control_pagos.control_pagos import actualizar_control_pagos


def actualizar_control_pagos_service(ruta_planilla, ruta_control, numero_armada):
    return actualizar_control_pagos(ruta_planilla, ruta_control, numero_armada)
