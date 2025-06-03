import os
import datetime
import pandas as pd
from docx import Document
from utils import *

def generar_conformidad_desde_excel(fila, plantilla_path, carpeta_base):
    docente = str(getattr(fila, "Docente", "N/A"))
    nombre_docente = limpiar_nombre_archivo(docente)
    carpeta_final = os.path.join(carpeta_base, "FASE FINAL", nombre_docente)
    os.makedirs(carpeta_final, exist_ok=True)

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

    clasif_cant_horas = clasif_valor / categoria_valor
    horas_clasif = f"{int(round(clasif_cant_horas))} horas de clasificación"

    if clasif_valor == 0:
        descripcion_final = f"{descripcion} y {horas_disenio}"
    else:
        descripcion_final = f"{descripcion}, {horas_disenio} y {horas_clasif}"

    monto_categoria = getattr(fila, "Categoria_monto", 0)
    monto_categoria_letras = monto_a_letras(monto_categoria)
    monto_total = getattr(fila, "Subtotal_pago", 0)
    try:
        monto_total = float(monto_total)
        monto_total_str = f"{monto_total:,.2f}"
    except Exception:
        monto_total_str = str(monto_total)
    monto_total_letras = monto_a_letras(monto_total)

    meses_es = {
        "January": "enero", "February": "febrero", "March": "marzo",
        "April": "abril", "May": "mayo", "June": "junio",
        "July": "julio", "August": "agosto", "September": "septiembre",
        "October": "octubre", "November": "noviembre", "December": "diciembre"
    }
    hoy = datetime.now()
    nombre_mes = meses_es[hoy.strftime("%B")]
    fecha_formateada = f"{hoy.day} de {nombre_mes} de {hoy.year}"

    doc = Document(plantilla_path)
    for p in doc.paragraphs:
        if "nombre" in p.text:
            p.text = p.text.replace("nombre", nombre_docente)
        if "numero_ruc" in p.text:
            p.text = p.text.replace("numero_ruc", ruc)
        if "descripcion" in p.text:
            p.text = p.text.replace("descripcion", descripcion_final)
        if "numero_orden" in p.text:
            p.text = p.text.replace("numero_orden", "______________")
        if "monto_subtotal" in p.text:
            p.text = p.text.replace("monto_subtotal", f"S/. {monto_total_str} ({monto_total_letras})")
        if "categoria_monto" in p.text:
            p.text = p.text.replace("categoria_monto", f"S/. {monto_categoria:.2f} ({monto_categoria_letras})")
        if "fecha" in p.text:
            p.text = p.text.replace("fecha", fecha_formateada)

    nombre_archivo = f"CONFORMIDAD - {nombre_docente} - {hoy.month} {hoy.year}.docx"
    output_path = os.path.join(carpeta_final, nombre_archivo)
    doc.save(output_path)
    return output_path

def procesar_planilla(excel_path, hoja, carpeta_base):
    plantilla_path = os.path.join("Modelos_documentos", "CONFORMIDAD - MODELO.docx")
    df = pd.read_excel(excel_path, sheet_name=hoja)
    generados, errores = [], []
    for _, fila in df.iterrows():
        try:
            output = generar_conformidad_desde_excel(fila, plantilla_path, carpeta_base)
            generados.append(os.path.basename(output))
        except Exception as e:
            errores.append((fila.get("Docente", "Sin nombre"), str(e)))
    return generados, errores