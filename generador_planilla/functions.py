import pandas as pd
import os
from fuzzywuzzy import process, fuzz
import unicodedata

# FUNCIONES PARA GENERACIÓN DE PLANILLA

def cargar_archivo(ruta):
    extension = os.path.splitext(ruta)[-1].lower()
    if extension == ".csv":
        try:
            return pd.read_csv(ruta, sep=',')
        except Exception:
            return pd.read_csv(ruta)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}")
    
def limpiar_docentes(df, col):
    df[col] = df[col].astype(str).str.strip()
    return df

def agrupar_y_calcular(df, datos_docentes, col_curso):
    agrupado = (
        df.groupby('docente')
          .agg(curso=(col_curso, lambda x: ' / '.join(x)),
               cantidad_cursos=(col_curso, 'count'))
          .reset_index()
    )
    nombres_base = datos_docentes['Docente'].tolist()
    agrupado['Docente'] = [
        process.extractOne(n, nombres_base, scorer=fuzz.token_sort_ratio)[0]
        if process.extractOne(n, nombres_base, scorer=fuzz.token_sort_ratio)[1] >= 85 else None
        for n in agrupado['docente']
    ]
    agrupado = agrupado.merge(
        datos_docentes[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N°. Ruc', 'Estado']],
        on='Docente', how='left'
    )
    agrupado['Curso Dictado'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 28
    agrupado['Diseño de Examenes'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 4
    return agrupado


def agregar_clasificacion(df, ruta_clasificacion, normalizar_texto):
    if os.path.exists(ruta_clasificacion):
        try:
            clasif_df = pd.read_excel(ruta_clasificacion, header=1)
            clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()
            clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)
            df['docente_norm'] = df['Docente'].apply(normalizar_texto)
            df = df.merge(clasif_df[['docente_norm', 'Monto']], on='docente_norm', how='left')
            df['Examen Clasif.'] = df['Monto'].fillna(0)
            df.drop(columns=['docente_norm', 'Monto'], inplace=True)
        except Exception as e:
            print(f"⚠️ Error al leer archivo de clasificación: {e}")
            df['Examen Clasif.'] = 0
    else:
        df['Examen Clasif.'] = 0
    return df

def construir_tabla(df):
    tabla = pd.DataFrame({
        'N°': range(1, len(df) + 1),
        'Docente': df['Docente'],
        'Sede': df['Sede'],
        'Categoria (Letra)': df['Categoria (Letra)'],
        'Categoria (Monto)': df['Categoria (Monto)'],
        'N°. Ruc': df['N°. Ruc'],
        'Curso': df['curso'],
        'Curso Dictado': df['Curso Dictado'],
        'Extra Curso': 0,
        'Cantidad Cursos': df['cantidad_cursos'],
        'Diseño de Examenes': df['Diseño de Examenes'],
        'Examen Clasif.': df['Examen Clasif.'],
        'Total Pago S/.': 0,
        'Estado': df['Estado']
    })
    tabla['Total Pago S/.'] = (
        tabla['Curso Dictado'] + tabla['Extra Curso'] +
        tabla['Diseño de Examenes'] + tabla['Examen Clasif.']
    )
    return tabla

def ajustar_nivel(row):
        nivel = row['nivel']
        idioma = row['idioma']
        ciclo = row['ciclo']
        if nivel == 'General':
            if idioma == 'Inglés':
                return 'Posgrado Intermedio'
            elif idioma == 'Portugués':
                try:
                    ciclo_num = int(str(ciclo).strip())
                except Exception:
                    ciclo_num = None
                if ciclo_num is not None:
                    if 1 <= ciclo_num <= 4:
                        return 'Posgrado Básico'
                    elif 5 <= ciclo_num <= 8:
                        return 'Posgrado Intermedio'
        return nivel

def ajustar_modalidad(row):
    dias = row['dias']
    hora_fin = row['horafin']
    modalidad = row['modalidad']

    if (hora_fin == '22:30:00'or hora_fin == '15:00:00') and (dias == '{MONDAY,TUESDAY,WEDNESDAY,THURSDAY}' or dias == '{SATURDAY,SUNDAY}'):
        modalidad = 'INTENSIVO VIRTUAL'

    return modalidad

def crear_df_carga(datos, estado_planilla):
    datos = datos.copy()

    datos['nivel'] = datos.apply(ajustar_nivel, axis=1)

    df = pd.DataFrame({
        'Dias': datos['dias'].apply(traducir_dias),
        'H. Inicio': datos['horainicio'].astype(str).str[:5],
        'H. Fin': datos['horafin'].astype(str).str[:5],
        'Idioma': datos['idioma'],
        'Nivel': datos['nivel'].str.replace('Ã¡', 'á', regex=False),
        'Ciclo': datos['ciclo'],
        'Curso': datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1),
        'Sede': datos['sede'],
        'Sec.': '',
        'Matr.': datos['matriculados'],
        'Docente': datos['docente'],
        'Modalidad': datos['modalidad'],
        'Estado Planilla': estado_planilla
    })
    df = df.sort_values(by='Docente').reset_index(drop=True)
    df.insert(0, 'N°', range(1, len(df) + 1))
    return df

def ordenar_hojas_excel(wb, hojas_ordenadas):
    hojas_existentes = wb.sheetnames
    nuevas_hojas = [hoja for hoja in hojas_ordenadas if hoja in hojas_existentes]
    for idx, hoja in enumerate(nuevas_hojas):
        wb._sheets.insert(idx, wb[hoja])
    wb._sheets = wb._sheets[:len(nuevas_hojas)] + [s for s in wb._sheets if s.title not in nuevas_hojas]
    return wb

def traducir_dias(dias_raw: str) -> str:
    dias_dict = {
        'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié',
        'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'
    }
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)

def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in texto if not unicodedata.combining(c)])