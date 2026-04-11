import os

import pandas as pd

from core.planillas.transformations import (
    ajustar_nivel,
    expandir_filas_carga_intensivos,
    expandir_texto_cursos_intensivos,
)


def construir_tabla_planilla(df, es_enero=False, monto_bono=0):
    df = df.copy()
    campos_requeridos = {
        'curso': '',
        'Curso Dictado': 0,
        'cantidad_cursos': 0,
        'Diseño de Examenes': 0,
        'Examen Clasif.': 0,
        'Servicio Actualización': 0,
    }

    for campo, valor_default in campos_requeridos.items():
        if campo not in df.columns:
            df[campo] = valor_default
        else:
            df[campo] = df[campo].fillna(valor_default)

    columnas_tabla = {
        'N°': range(1, len(df) + 1),
        'Docente': df['Docente'],
        'Sede': df['Sede'],
        'Categoria (Letra)': df['Categoria (Letra)'],
        'Categoria (Monto)': df['Categoria (Monto)'],
        'N_Ruc': df['N_Ruc'],
        'Curso': df['curso'],
        'Curso Dictado': df['Curso Dictado'],
    }

    if es_enero:
        columnas_tabla['Bono'] = monto_bono

    columnas_tabla.update({
        'Extra Curso': 0,
        'Cantidad Cursos': df['cantidad_cursos'],
        'Diseño de Examenes': df['Diseño de Examenes'],
        'Examen Clasif.': df['Examen Clasif.'],
        'Servicio Actualización': df['Servicio Actualización'],
        'Total Pago S/.': 0,
        'Estado': df['Estado'],
    })

    tabla = pd.DataFrame(columnas_tabla)
    tabla['Curso'] = tabla['Curso'].apply(expandir_texto_cursos_intensivos)

    def agregar_servicio_a_curso(row):
        curso_actual = str(row['Curso']).strip()
        servicio_actualizacion = row['Servicio Actualización']

        if servicio_actualizacion > 0:
            texto_servicio = 'Servicio de actualización de materiales de enseñanza'
            if curso_actual and curso_actual != '' and curso_actual != 'nan':
                return f"{curso_actual} / {texto_servicio}"
            return texto_servicio

        return curso_actual

    tabla['Curso'] = tabla.apply(agregar_servicio_a_curso, axis=1)

    if es_enero:
        tabla['Total Pago S/.'] = (
            tabla['Curso Dictado'] + tabla['Bono'] + tabla['Extra Curso'] +
            tabla['Diseño de Examenes'] + tabla['Examen Clasif.'] +
            tabla['Servicio Actualización']
        )
    else:
        tabla['Total Pago S/.'] = (
            tabla['Curso Dictado'] + tabla['Extra Curso'] + tabla['Diseño de Examenes'] +
            tabla['Examen Clasif.'] + tabla['Servicio Actualización']
        )

    tabla = tabla.sort_values('Docente').reset_index(drop=True)
    tabla['N°'] = range(1, len(tabla) + 1)

    return tabla


def construir_tabla_carga_academica(datos, estado_planilla, traducir_dias_fn):
    datos = expandir_filas_carga_intensivos(datos)

    if 'nivel' not in datos.columns:
        datos['nivel'] = datos.apply(ajustar_nivel, axis=1)
    if 'Curso' not in datos.columns:
        datos['Curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

    df = pd.DataFrame({
        'Dias': datos['dias'].apply(traducir_dias_fn),
        'H. Inicio': datos['horainicio'].astype(str).str[:5],
        'H. Fin': datos['horafin'].astype(str).str[:5],
        'Idioma': datos['idioma'],
        'Nivel': datos['nivel'].str.replace('Ã¡', 'á', regex=False),
        'Ciclo': datos['ciclo'],
        'Curso': datos['Curso'],
        'Sede': datos['sede'],
        'Sec.': '',
        'Matr.': datos['matriculados'],
        'Docente': datos['docente'],
        'Modalidad': datos['modalidad'],
        'Estado Planilla': estado_planilla,
    })
    df = df.sort_values(by='Docente').reset_index(drop=True)
    df.insert(0, 'N°', range(1, len(df) + 1))
    return df


def construir_tabla_coordinacion(
    ruta_coordinacion,
    normalizar_texto,
    datos_docentes,
    cargar_excel_con_cache_fn,
    preparar_coordinacion_agrupada_fn,
):
    if not os.path.exists(ruta_coordinacion):
        return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])

    try:
        coordinacion_df = cargar_excel_con_cache_fn(ruta_coordinacion, sheet_name=0, header=0)

        coordinacion_agrupado = preparar_coordinacion_agrupada_fn(coordinacion_df, normalizar_texto)
        if coordinacion_agrupado is None:
            return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])

        datos_docentes_temp = datos_docentes.copy()
        datos_docentes_temp['docente_norm'] = datos_docentes_temp['Docente'].apply(normalizar_texto)

        coordinacion_con_categoria = coordinacion_agrupado.merge(
            datos_docentes_temp[['docente_norm', 'Docente', 'Categoria (Letra)', 'Categoria (Monto)']],
            on='docente_norm', how='left'
        )

        coordinacion_con_categoria['Docente_Final'] = coordinacion_con_categoria['Docente'].fillna(
            coordinacion_con_categoria['Docente_Original']
        )
        coordinacion_con_categoria['Categoria (Letra)'] = coordinacion_con_categoria['Categoria (Letra)'].fillna('N/A')
        coordinacion_con_categoria['Categoria (Monto)'] = coordinacion_con_categoria['Categoria (Monto)'].fillna(0)

        coordinacion_con_categoria['Monto_Total'] = (
            coordinacion_con_categoria['Horas_Total'] *
            coordinacion_con_categoria['Categoria (Monto)']
        )

        tabla_coordinacion = pd.DataFrame({
            'N°': range(1, len(coordinacion_con_categoria) + 1),
            'Docente': coordinacion_con_categoria['Docente_Final'],
            'Categoría (Letra)': coordinacion_con_categoria['Categoria (Letra)'],
            'Categoría (Monto)': coordinacion_con_categoria['Categoria (Monto)'],
            'Horas Totales': coordinacion_con_categoria['Horas_Total'],
            'Monto Total': coordinacion_con_categoria['Monto_Total'],
        })

        tabla_coordinacion = tabla_coordinacion.sort_values('Docente').reset_index(drop=True)
        tabla_coordinacion['N°'] = range(1, len(tabla_coordinacion) + 1)

        print(f"✅ Tabla de coordinación creada con {len(tabla_coordinacion)} docentes")
        return tabla_coordinacion

    except Exception as e:
        print(f"⚠️ Error al crear tabla de coordinación: {e}")
        return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])
