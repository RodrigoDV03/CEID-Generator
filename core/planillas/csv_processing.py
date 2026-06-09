import os
import re
import pandas as pd


COLUMNAS_PLANILLA_CANONICAS = {
    'Programa Educativo': 'programa_educativo',
    'Idioma': 'idioma',
    'Nivel': 'nivel',
    'Ciclo': 'ciclo',
    'Modalidad': 'modalidad',
    'Sede': 'sede',
    'Dia': 'dias',
    'Día': 'dias',
    'Hora Inicio': 'horainicio',
    'Hora Fin': 'horafin',
    'Docente': 'docente',
    'Fecha Inicio Clases': 'fecha_inicio_clases',
    'Fecha Fin Clases': 'fecha_fin_clases',
    'Inicio Matricula': 'inicio_matricula',
    'Cierre Matricula': 'cierre_matricula',
    'Capacidad Min': 'capacidad_min',
    'Capacidad Max': 'capacidad_max',
    'Total Matriculados': 'matriculados',
}


def normalizar_columnas_planilla(df):
    df_normalizado = df.copy()
    df_normalizado.columns = [str(col).strip() for col in df_normalizado.columns]

    renombres = {
        origen: destino
        for origen, destino in COLUMNAS_PLANILLA_CANONICAS.items()
        if origen in df_normalizado.columns and destino not in df_normalizado.columns
    }

    if renombres:
        df_normalizado = df_normalizado.rename(columns=renombres)

    if 'Curso' not in df_normalizado.columns and {'idioma', 'nivel', 'ciclo'}.issubset(df_normalizado.columns):
        df_normalizado['Curso'] = (
            df_normalizado['idioma'].astype(str).str.strip() + ' ' +
            df_normalizado['nivel'].astype(str).str.strip() + ' ' +
            df_normalizado['ciclo'].astype(str).str.strip()
        )

    return df_normalizado


def cargar_archivo(ruta):
    extension = os.path.splitext(ruta)[-1].lower()
    if extension == ".csv":
        # Intentar lectura normal primero con deteccion automatica de delimitador.
        try:
            # Primero intentar con punto y coma (;)
            df = pd.read_csv(ruta, sep=';', encoding='utf-8-sig')
            if len(df.columns) > 1:
                return df

            # Si solo tiene una columna, intentar con coma
            df = pd.read_csv(ruta, sep=',', encoding='utf-8-sig')
            if len(df.columns) > 1:
                return df

            # Si todavia no funciona, usar el parser personalizado
            return parsear_csv_comillas_dobles(ruta)
        except Exception as e:
            print(f"⚠️ Error al leer CSV con pandas: {e}")
            return parsear_csv_comillas_dobles(ruta)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}")


def parsear_csv_comillas_dobles(ruta):
    with open(ruta, 'r', encoding='utf-8') as f:
        lineas = f.readlines()

    datos_parseados = []

    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue

        # Remover comillas externas si existen
        if linea.startswith('"') and linea.endswith('"'):
            linea = linea[1:-1]

        # Dividir la linea considerando el formato especial
        campos = []
        campo_actual = ""
        dentro_comillas = False
        dentro_llaves = 0
        i = 0

        while i < len(linea):
            char = linea[i]

            # Manejar comillas dobles ""
            if i < len(linea) - 1 and linea[i:i + 2] == '""':
                if not dentro_comillas:
                    dentro_comillas = True
                    i += 2
                    continue
                else:
                    dentro_comillas = False
                    i += 2
                    continue

            # Manejar llaves para dias
            elif char == '{':
                dentro_llaves += 1
                campo_actual += char
            elif char == '}':
                dentro_llaves -= 1
                campo_actual += char

            # Manejar comas separadoras
            elif char == ',' and not dentro_comillas and dentro_llaves == 0:
                campos.append(campo_actual.strip())
                campo_actual = ""
            else:
                campo_actual += char

            i += 1

        # Agregar el ultimo campo
        if campo_actual:
            campos.append(campo_actual.strip())

        # Limitar a 13 columnas
        if len(campos) > 13:
            campos = campos[:13]

        datos_parseados.append(campos)

    if datos_parseados:
        headers = datos_parseados[0]
        filas = datos_parseados[1:]

        # Verificar que todas las filas tengan el mismo numero de columnas que el header
        max_cols = len(headers)
        filas_ajustadas = []
        for fila in filas:
            if len(fila) < max_cols:
                fila = fila + [''] * (max_cols - len(fila))
            elif len(fila) > max_cols:
                fila = fila[:max_cols]
            filas_ajustadas.append(fila)

        print(f"✅ Creando DataFrame con {len(headers)} columnas y {len(filas_ajustadas)} filas")
        return pd.DataFrame(filas_ajustadas, columns=headers)

    return pd.DataFrame()


def limpiar_docentes(df, col):
    df[col] = df[col].astype(str).str.strip()
    return df


def docente_es_valido(valor):
    if pd.isna(valor):
        return False

    texto = str(valor).strip()
    if not texto:
        return False

    return texto.lower() not in {'nan', 'none', 'null'} and texto != ','


def filtrar_docentes_validos(df, col='docente'):
    if col not in df.columns:
        return df.copy()

    return df[df[col].apply(docente_es_valido)].copy()


def procesar_csv_nuevo_formato(df):
    df_procesado = df.copy()

    df_procesado = normalizar_columnas_planilla(df_procesado)

    # Mapeo directo de columnas
    mapeo_columnas = {
        'Modalidad': 'modalidad',
        'Sede': 'sede',
        'Docente': 'docente',
        'Total Matriculados': 'matriculados'
    }

    # Renombrar columnas directas
    for col_origen, col_destino in mapeo_columnas.items():
        if col_origen in df_procesado.columns:
            df_procesado[col_destino] = df_procesado[col_origen]

    # Procesar "Detalle Curso" para extraer idioma, nivel y ciclo
    if 'Detalle Curso' in df_procesado.columns:
        def extraer_info_curso(detalle):
            if pd.isna(detalle):
                return pd.Series({'idioma': '', 'nivel': '', 'ciclo': ''})

            detalle_str = str(detalle).strip()
            patron = r'^([A-Za-zÀ-ÿ]+)\s+([A-Za-zÀ-ÿ]+)\s+(\d+)$'
            match = re.match(patron, detalle_str)

            if match:
                idioma = match.group(1)
                nivel = match.group(2)
                ciclo = match.group(3)
                return pd.Series({'idioma': idioma, 'nivel': nivel, 'ciclo': ciclo})
            else:
                partes = detalle_str.split()
                if len(partes) >= 3:
                    ciclo = partes[-1]
                    nivel = partes[-2]
                    idioma = ' '.join(partes[:-2])
                    return pd.Series({'idioma': idioma, 'nivel': nivel, 'ciclo': ciclo})
                else:
                    return pd.Series({'idioma': detalle_str, 'nivel': '', 'ciclo': ''})

        df_info = df_procesado['Detalle Curso'].apply(extraer_info_curso)
        df_procesado['idioma'] = df_info['idioma']
        df_procesado['nivel'] = df_info['nivel']
        df_procesado['ciclo'] = df_info['ciclo']

    # Procesar "Horario Completo" para extraer dias, horainicio y horafin
    if 'Horario Completo' in df_procesado.columns:
        def extraer_info_horario(horario):
            if pd.isna(horario):
                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})

            horario_str = str(horario).strip()

            dias_map = {
                'Lun': 'MONDAY',
                'Mar': 'TUESDAY',
                'Mie': 'WEDNESDAY',
                'Mié': 'WEDNESDAY',
                'Jue': 'THURSDAY',
                'Vie': 'FRIDAY',
                'Sab': 'SATURDAY',
                'Sáb': 'SATURDAY',
                'Dom': 'SUNDAY'
            }

            try:
                if ' ' in horario_str:
                    partes = horario_str.split(' ', 1)
                    dias_str = partes[0]
                    horas_str = partes[1] if len(partes) > 1 else ''

                    dias_lista = [d.strip() for d in dias_str.split(',')]
                    dias_ingles = [dias_map.get(d, d) for d in dias_lista]
                    dias_formateado = '{' + ','.join(dias_ingles) + '}'

                    if ' - ' in horas_str:
                        horas_partes = horas_str.split(' - ')
                        hora_inicio_str = horas_partes[0].strip()
                        hora_fin_str = horas_partes[1].strip() if len(horas_partes) > 1 else ''

                        def convertir_12h_a_24h(hora_12h):
                            try:
                                import datetime
                                hora_obj = datetime.datetime.strptime(hora_12h, '%I:%M %p')
                                return hora_obj.strftime('%H:%M:%S')
                            except Exception:
                                return hora_12h

                        hora_inicio = convertir_12h_a_24h(hora_inicio_str)
                        hora_fin = convertir_12h_a_24h(hora_fin_str)

                        return pd.Series({
                            'dias': dias_formateado,
                            'horainicio': hora_inicio,
                            'horafin': hora_fin
                        })

                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})
            except Exception as e:
                print(f"Error al procesar horario '{horario_str}': {e}")
                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})

        df_horario = df_procesado['Horario Completo'].apply(extraer_info_horario)
        df_procesado['dias'] = df_horario['dias']
        df_procesado['horainicio'] = df_horario['horainicio']
        df_procesado['horafin'] = df_horario['horafin']

    return df_procesado
