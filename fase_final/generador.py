import os
import pandas as pd
from docx import Document
from utils import *

def generar_conformidad_desde_excel(fila, plantilla_path, ruta_salida):

    docente = str(getattr(fila, "Docente", "N/A"))
    nombre_docente = limpiar_nombre_archivo(docente)

    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla: {plantilla_path}")

    ruc = limpiar_numero(getattr(fila, "N_Ruc", ""))
    descripcion_raw = str(getattr(fila, "Curso", ""))
    descripcion = redactar_cursos(descripcion_raw)
    cant_cursos = int(getattr(fila, "Cantidad_cursos", 0))
    disenio_cant_horas = cant_cursos * 4
    horas_disenio = f"{int(round(disenio_cant_horas))} horas de diseño de exámenes"
    clasif_valor = int(getattr(fila, "Examen_clasif", 0))

    categoria_valor = getattr(fila, "Categoria_monto", 1)
    if pd.isna(categoria_valor):
        categoria_valor = 1
    else:
        categoria_valor = int(categoria_valor)

    clasif_cant_horas = clasif_valor / categoria_valor if categoria_valor else 0
    horas_clasif = f"{int(round(clasif_cant_horas))} horas de clasificación"

    if clasif_valor == 0:
        descripcion_final = f"{descripcion} y {horas_disenio}"
    else:
        descripcion_final = f"{descripcion}, {horas_disenio} y {horas_clasif}"

    monto_categoria = getattr(fila, "Categoria_monto", 0)
    monto_categoria_letras = monto_a_letras(monto_categoria)
    monto_total = getattr(fila, "Subtotal_pago", 0)
    nro_contrato = getattr(fila, "Nro_Contrato", "")
    try:
        monto_total = float(monto_total)
        monto_total_str = f"{monto_total:,.2f}"
    except Exception:
        monto_total_str = str(monto_total)
    monto_total_letras = monto_a_letras(monto_total)

    doc = Document(plantilla_path)
    for p in doc.paragraphs:
        for run in p.runs:
            if "nombre" in run.text:
                run.text = run.text.replace("nombre", str(nombre_docente))
            if "ruc" in run.text:
                run.text = run.text.replace("ruc", str(ruc))
            if "descripcion_cursos" in run.text:
                run.text = run.text.replace("descripcion_cursos", str(descripcion_final))
            if "numero_orden" in run.text:
                run.text = run.text.replace("numero_orden", "______________")
            if "monto_subtotal" in run.text:
                run.text = run.text.replace("monto_subtotal", f"S/. {monto_total_str} ({str(monto_total_letras)})")
            if "monto_hora" in run.text:
                run.text = run.text.replace("monto_hora", f"S/. {monto_categoria:.2f} ({str(monto_categoria_letras)})")
            if "Nro_Contrato" in run.text:
                run.text = run.text.replace("Nro_Contrato", str(nro_contrato))

    carpeta_final = os.path.dirname(ruta_salida)
    os.makedirs(carpeta_final, exist_ok=True)

    try:
        doc.save(ruta_salida)
    except Exception as e:
        print(f"Error al guardar {ruta_salida}: {e}")
        raise
    return ruta_salida


def procesar_planilla(ruta_excel, hoja, carpeta_salida, mes, año):
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    generados = []
    errores = []

    for _, fila in df.iterrows():
        try:
            estado = fila["ESTADO"].strip().lower()
            nombre = fila["Docente"].strip()
            nombre_docente = limpiar_nombre_archivo(nombre)

            if estado == "contrato":
                plantilla = "Modelos_documentos/CONFORMIDAD CONTRATO - MODELO.docx"
            elif estado == "tercero":
                plantilla = "Modelos_documentos/CONFORMIDAD TERCERO - MODELO.docx"
            else:
                raise ValueError(f"Estado inválido: {estado}")

            carpeta_final = os.path.join(carpeta_salida, "FASE FINAL", nombre_docente)
            os.makedirs(carpeta_final, exist_ok=True)

            nombre_archivo = f"CONFORMIDAD - {nombre} - {mes} {año}.docx"
            ruta_salida = os.path.join(carpeta_final, nombre_archivo)

            generar_conformidad_desde_excel(fila, plantilla, ruta_salida)
            generados.append(os.path.join("FASE FINAL", nombre_docente, nombre_archivo))
        except Exception as e:
            errores.append((fila.get("Docente", "Desconocido"), str(e)))

    return generados, errores