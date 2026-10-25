"""
Nombres de columna de las hojas de salida de la planilla.

Antes estos nombres vivían como strings sueltos, repetidos (y a veces
ligeramente distintos) en varios archivos: `functions.py`,
`generador_planilla.py`, `table_builders.py`,
`core/fases/services/excel_reader_service.py`. Con el tiempo eso hizo que
"lo mismo" tuviera varias grafías conviviendo — por ejemplo 'Disenio_examenes'
(con la falta de tilde ya arrastrada como si fuera el nombre real) frente a
'Diseño de Examenes'. Este módulo es ahora la única fuente de verdad.

Conviven dos familias de nombres para el mismo concepto, porque hay dos
formatos de hoja en juego:

- Los nombres "internos" son los que se usan mientras se arma la tabla con
  pandas, y también los que ve el personal en la hoja legible
  "Planilla <mes>".
- Los nombres "RESUMIDA_*" son el formato snake_case de la hoja
  "Planilla_Generador", que además lee
  `core/fases/services/excel_reader_service.py` para generar los documentos
  de fase. Ese nombre es un contrato externo (ya lo consume otro código) y
  no se cambia acá para no romperlo — solo se centraliza para que no vuelva
  a escribirse distinto en cada archivo.
"""

# ----- Identificación del docente -----
DOCENTE = 'Docente'
DOCENTE_ORIGINAL = 'docente'  # tal como viene en el CSV de cursos, antes de emparejar con la base de docentes
ID_PROGRAMACION = 'id_programacion'  # columna "ID Programación" del CSV de cursos
N_RUC = 'N_Ruc'
N_RUC_DISPLAY = 'N.° RUC'  # nombre de columna en la hoja "Planilla <mes>", legible para el personal
SEDE = 'Sede'
CATEGORIA_LETRA = 'Categoria (Letra)'
CATEGORIA_MONTO = 'Categoria (Monto)'
ESTADO = 'Estado'

# Nombres alternativos que puede traer la columna de RUC en el Excel de
# docentes que sube el personal (no todos usan el mismo formato).
ALIAS_COLUMNA_RUC = ('N_Ruc', 'N. Ruc', 'N.° RUC')

# Nombres alternativos que puede traer la columna de número de contrato,
# tanto en el Excel de docentes como en el de control de pagos.
ALIAS_COLUMNA_CONTRATO = (
    'N.° Contrato',
    'N° Contrato',
    'Nro_Contrato',
    'Nro_contrato',
    'Numero de contrato',
    'Número de contrato',
)

# ----- Montos por tipo de servicio (nombre interno, mientras se calcula) -----
CURSO_DICTADO = 'Curso Dictado'
DISENO_EXAMENES = 'Diseño de Examenes'
EXAMEN_CLASIFICACION = 'Examen Clasif.'
SERVICIO_ACTUALIZACION = 'Servicio Actualización'
CANTIDAD_CURSOS = 'cantidad_cursos'
TOTAL_PAGO = 'Total Pago S/.'

# ----- Identificadores de tipo de servicio (hoja expandida "Planilla_Generador") -----
TIPO_SERVICIO_CURSO_DICTADO = 'CURSO_DICTADO'
TIPO_SERVICIO_DISENO_EXAMENES = 'DISENO_EXAMENES'
TIPO_SERVICIO_EXAMEN_CLASIF = 'EXAMEN_CLASIF'
TIPO_SERVICIO_ACTUALIZACION = 'SERVICIO_ACTUALIZACION'

NOMBRES_TIPO_SERVICIO = {
    TIPO_SERVICIO_CURSO_DICTADO: 'Curso Académico',
    TIPO_SERVICIO_DISENO_EXAMENES: 'Diseño de Exámenes',
    TIPO_SERVICIO_EXAMEN_CLASIF: 'Examen de Clasificación',
    TIPO_SERVICIO_ACTUALIZACION: 'Servicio de Actualización',
}

ORDEN_TIPO_SERVICIO = {
    TIPO_SERVICIO_CURSO_DICTADO: 1,
    TIPO_SERVICIO_DISENO_EXAMENES: 2,
    TIPO_SERVICIO_EXAMEN_CLASIF: 3,
    TIPO_SERVICIO_ACTUALIZACION: 4,
}

# ----- Nombres snake_case de la hoja "Planilla_Generador" (contrato externo) -----
RESUMIDA_CATEGORIA_LETRA = 'Categoria_letra'
RESUMIDA_CATEGORIA_MONTO = 'Categoria_monto'
RESUMIDA_DISENO_EXAMENES = 'Disenio_examenes'
RESUMIDA_EXAMEN_CLASIF = 'Examen_clasif'
RESUMIDA_TOTAL_PAGO = 'Total_pago'
RESUMIDA_ESTADO_DOCENTE = 'Estado_docente'
RESUMIDA_DOCENTE_IDIOMA = 'Docente_idioma'
RESUMIDA_TIPO_DOCUMENTO = 'Tipo_documento'
RESUMIDA_NUMERO_DNI = 'Numero_dni'
RESUMIDA_NUMERO_CELULAR = 'Numero_celular'
RESUMIDA_DOMICILIO_DOCENTE = 'Domicilio_docente'
RESUMIDA_CORREO_PERSONAL = 'Correo_personal'
RESUMIDA_NRO_CONTRATO = 'Nro_Contrato'
RESUMIDA_SERVICIO_ACTUALIZACION = 'Servicio_actualizacion'
RESUMIDA_HORAS_TOTAL = 'Horas_Total'

# Columnas "extra" que se traen del Excel de docentes tal cual, antes de
# renombrarlas al formato snake_case de la hoja Planilla_Generador.
COLUMNAS_EXTRA_DOCENTES = [
    'Idioma', 'Tipo Documento', 'Nro. Documento', 'Celular', 'Dirección',
    'Correo personal', 'N.° Contrato',
]

# Mapeo interno -> snake_case, usado al construir la hoja "Planilla_Generador".
RENOMBRES_TABLA_GENERADOR = {
    CATEGORIA_LETRA: RESUMIDA_CATEGORIA_LETRA,
    CATEGORIA_MONTO: RESUMIDA_CATEGORIA_MONTO,
    DISENO_EXAMENES: RESUMIDA_DISENO_EXAMENES,
    EXAMEN_CLASIFICACION: RESUMIDA_EXAMEN_CLASIF,
    TOTAL_PAGO: RESUMIDA_TOTAL_PAGO,
    ESTADO: RESUMIDA_ESTADO_DOCENTE,
    'Idioma': RESUMIDA_DOCENTE_IDIOMA,
    'Tipo Documento': RESUMIDA_TIPO_DOCUMENTO,
    'Nro. Documento': RESUMIDA_NUMERO_DNI,
    'Celular': RESUMIDA_NUMERO_CELULAR,
    'Dirección': RESUMIDA_DOMICILIO_DOCENTE,
    'Correo personal': RESUMIDA_CORREO_PERSONAL,
    'N.° Contrato': RESUMIDA_NRO_CONTRATO,
    SERVICIO_ACTUALIZACION: RESUMIDA_SERVICIO_ACTUALIZACION,
}

# Columnas de la hoja "Planilla_Generador" resumida (una fila por docente).
COLUMNAS_PLANILLA_GENERADOR_RESUMIDA = [
    'N°', DOCENTE, N_RUC, RESUMIDA_CATEGORIA_LETRA, RESUMIDA_CATEGORIA_MONTO, SEDE,
    'Curso_Virtual', 'Curso_Presencial', CANTIDAD_CURSOS, CURSO_DICTADO,
    RESUMIDA_DISENO_EXAMENES, RESUMIDA_EXAMEN_CLASIF, RESUMIDA_HORAS_TOTAL, RESUMIDA_SERVICIO_ACTUALIZACION,
    RESUMIDA_TOTAL_PAGO, RESUMIDA_ESTADO_DOCENTE, RESUMIDA_DOCENTE_IDIOMA, RESUMIDA_TIPO_DOCUMENTO,
    RESUMIDA_NUMERO_DNI, RESUMIDA_NUMERO_CELULAR, RESUMIDA_DOMICILIO_DOCENTE, RESUMIDA_CORREO_PERSONAL,
    RESUMIDA_NRO_CONTRATO,
]

# Subconjunto de texto vs. numéricas dentro de esas mismas columnas (para
# rellenar valores por defecto correctos si faltan).
COLUMNAS_RESUMIDA_TEXTO = [
    N_RUC, RESUMIDA_CATEGORIA_LETRA, SEDE, RESUMIDA_ESTADO_DOCENTE, RESUMIDA_DOCENTE_IDIOMA, RESUMIDA_TIPO_DOCUMENTO,
    RESUMIDA_NUMERO_DNI, RESUMIDA_NUMERO_CELULAR, RESUMIDA_DOMICILIO_DOCENTE, RESUMIDA_CORREO_PERSONAL, RESUMIDA_NRO_CONTRATO,
]
COLUMNAS_RESUMIDA_NUMERICAS = [
    RESUMIDA_CATEGORIA_MONTO, CANTIDAD_CURSOS, CURSO_DICTADO, RESUMIDA_DISENO_EXAMENES,
    RESUMIDA_EXAMEN_CLASIF, RESUMIDA_HORAS_TOTAL, RESUMIDA_SERVICIO_ACTUALIZACION, RESUMIDA_TOTAL_PAGO,
]

# Columnas de la hoja "Planilla_Generador" expandida (una fila por servicio).
COLUMNAS_EXPANDIDA_PRINCIPALES = [
    'N°', DOCENTE, 'Servicio_Nro', 'Tipo_Servicio_Desc', 'Curso_Individual',
    'Modalidad_Curso', 'Horas_Servicio', 'Monto_Individual',
]
COLUMNAS_EXPANDIDA_RESUMEN = [
    CANTIDAD_CURSOS, CURSO_DICTADO, RESUMIDA_DISENO_EXAMENES, RESUMIDA_EXAMEN_CLASIF, RESUMIDA_SERVICIO_ACTUALIZACION,
]
COLUMNAS_EXPANDIDA_ADMINISTRATIVAS = [
    SEDE, RESUMIDA_CATEGORIA_LETRA, RESUMIDA_CATEGORIA_MONTO, N_RUC, RESUMIDA_ESTADO_DOCENTE, RESUMIDA_DOCENTE_IDIOMA,
    RESUMIDA_NUMERO_DNI, RESUMIDA_NUMERO_CELULAR, RESUMIDA_DOMICILIO_DOCENTE, RESUMIDA_CORREO_PERSONAL, RESUMIDA_NRO_CONTRATO,
]

# ----- Hoja "Servicio actualización" (construida a partir del archivo de coordinación) -----
COORD_CATEGORIA_LETRA = CATEGORIA_LETRA
COORD_CATEGORIA_MONTO = CATEGORIA_MONTO
COORD_HORAS_TOTALES = 'Horas Totales'
COORD_MONTO_TOTAL = 'Monto Total'
COLUMNAS_TABLA_COORDINACION = ['N°', DOCENTE, COORD_CATEGORIA_LETRA, COORD_CATEGORIA_MONTO, COORD_HORAS_TOTALES, COORD_MONTO_TOTAL]
