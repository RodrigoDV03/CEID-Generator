import os
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from docx import Document
from utils.functions import *

def generador_conformidad(fila, ruta_conformidad, ruta_destino, numero_armada):

    docente = str(getattr(fila, "Docente", "N/A"))
    nombre_docente = docente

    if not os.path.exists(ruta_conformidad):
        raise FileNotFoundError(f"No se encontró la plantilla: {ruta_conformidad}")

    ruc = limpiar_numero(getattr(fila, "N_Ruc", ""))
    descripcion_raw = str(getattr(fila, "Curso", ""))
    descripcion = redactar_cursos(descripcion_raw)
    cant_cursos = int(getattr(fila, "Cantidad_cursos", 0))
    disenio_cant_horas = cant_cursos * 4
    horas_disenio = f"{int(round(disenio_cant_horas))} horas de diseño de exámenes"
    clasif_valor = int(getattr(fila, "Examen_clasif", 0))

    monto_categoria_raw = getattr(fila, "Categoria_monto", 1)
    if pd.isna(monto_categoria_raw):
        monto_categoria = Decimal("1.00")
    else:
        monto_categoria = Decimal(str(monto_categoria_raw)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    clasif_cant_horas = clasif_valor / monto_categoria if monto_categoria else 0

    if clasif_cant_horas == 1:
        horas_clasif = f"{int(round(clasif_cant_horas))} hora de examen de clasificación"
    else:
        horas_clasif = f"{int(round(clasif_cant_horas))} horas de examen de clasificación"

    if clasif_valor == 0:
        descripcion_final = f"{descripcion} y {horas_disenio}"
    else:
        descripcion_final = f"{descripcion}, {horas_disenio} y {horas_clasif}"

    monto_categoria_letras = monto_a_letras(monto_categoria)
    monto_total = getattr(fila, "Subtotal_pago", 0)
    try:
        monto_total = float(monto_total)
        monto_total_str = f"{monto_total:,.2f}"
    except Exception:
        monto_total_str = str(monto_total)
    monto_total_letras = monto_a_letras(monto_total)

    nro_contrato = getattr(fila, "Nro_Contrato", "")
    try:
        nro_contrato = int(float(nro_contrato))
    except (ValueError, TypeError):
        pass

    doc = Document(ruta_conformidad)
    reemplazos_oficio = {
        "nombre": str(nombre_docente),
        "ruc": str(ruc),
        "descripcion_cursos": str(descripcion_final),
        "monto_subtotal": f"S/. {monto_total_str} ({str(monto_total_letras)})",
        "monto_hora": f"S/. {monto_categoria:.2f} ({str(monto_categoria_letras)})",
        "Nro_Contrato": str(nro_contrato),
        "numero_armada": str(numero_armada)
    }
    reemplazar_en_tablas(doc, reemplazos_oficio)
    reemplazar_en_parrafos(doc, reemplazos_oficio)

    carpeta_final = os.path.dirname(ruta_destino)
    os.makedirs(carpeta_final, exist_ok=True)

    try:
        doc.save(ruta_destino)
    except Exception as e:
        print(f"Error al guardar {ruta_destino}: {e}")
        raise
    return ruta_destino

def generar_control_avance(fila_control, doc_control, ruta_destino):

    nombre_docente = str(getattr(fila_control, "APELLIDOS Y NOMBRES", ""))

    idioma_docente = str(getattr(fila_control, "Especialidad", ""))
    numero_contrato = getattr(fila_control, "Numero de contrato", "")
    try:
        numero_contrato_str = str(int(float(numero_contrato)))
    except (ValueError, TypeError):
        numero_contrato_str = str(numero_contrato)
    monto_total = getattr(fila_control, "MONTO TOTAL PARA CONTRATO S/", 0)
    primera_armada = getattr(fila_control, "Primera armada", 0)
    segunda_armada = getattr(fila_control, "Segunda armada", 0)
    total_primera = getattr(fila_control, "Primera armada", 0)
    total_segunda = getattr(fila_control, "Segunda armada", 0)
    total_tercera = getattr(fila_control, "Tercera armada", 0)
    saldo_restante = getattr(fila_control, "Saldo restante", 0)

    saldo_primera = monto_total - primera_armada
    saldo_segunda = saldo_primera - segunda_armada

    # === Reemplazar en documento ===
    doc = Document(doc_control)
    reemplazos_armada = {
        "Nombre_Docente": str(nombre_docente),
        "Nro_Contrato": numero_contrato_str,
        "Idioma_Docente": str(idioma_docente),
        "Monto_Total": formato_soles(monto_total),
        "Total_Primera": formato_soles(total_primera),
        "Total_Segunda": formato_soles(total_segunda),
        "Total_Tercera": formato_soles(total_tercera),
        "Saldo_Restante": formato_soles(saldo_restante),
        "Saldo_Primera": formato_soles(saldo_primera),
        "Saldo_Segunda": formato_soles(saldo_segunda),
    }

    reemplazar_en_parrafos(doc, reemplazos_armada)
    reemplazar_en_tablas(doc, reemplazos_armada)

    carpeta_final = os.path.dirname(ruta_destino)
    os.makedirs(carpeta_final, exist_ok=True)

    # === Guardar archivo ===
    try:
        doc.save(ruta_destino)
    except Exception as e:
        print(f"Error al guardar {ruta_destino}: {e}")
        raise

    return ruta_destino


def procesar_planilla_fase_final(ruta_planilla, excel_control_pagos, hoja, carpeta_salida, mes, año, numero_armada):
    df = pd.read_excel(ruta_planilla, sheet_name=hoja)
    df_control = pd.read_excel(excel_control_pagos, sheet_name=0, header=1)

    # ============ GENERACIÓN DE CONFORMIDADES ============

    for _, fila in df.iterrows():
        try:
            docente = str(getattr(fila, "Docente", "N/A")).strip()
            estado = str(getattr(fila, "Estado_docente", "")).strip().upper()

            if estado == "CONTRATO" or estado == "Contrato":
                ruta_conformidad = ruta_absoluta_relativa("Modelos_documentos/conformidad_contrato.docx")
            elif estado == "TERCERO" or estado == "Tercero":
                ruta_conformidad = ruta_absoluta_relativa("Modelos_documentos/conformidad_tercero.docx")
            else:
                raise ValueError(f"Estado inválido: {estado}")

            carpeta_final = os.path.join(carpeta_salida, "FASE FINAL", docente)
            os.makedirs(carpeta_final, exist_ok=True)

            # === GENERAR DOCUMENTO DE CONFORMIDAD ===
            nombre_archivo_conformidad = f"CONFORMIDAD - {docente} - {mes} {año}.docx"
            ruta_destino_conformidad = os.path.join(carpeta_final, nombre_archivo_conformidad)
            generador_conformidad(fila, ruta_conformidad, ruta_destino_conformidad, numero_armada)

            print(f"{docente} - Documento de conformidad generado correctamente.")
        except Exception as e:
            print(f"Error procesando fila para {docente}: {e}")
            continue


    # ============ GENERACIÓN DE CONTROLES DE PAGO ============


    for _, fila_control in df_control.iterrows():
        try:
            docente = str(getattr(fila_control, "APELLIDOS Y NOMBRES", "N/A")).strip()

            if numero_armada == 'primera':
                doc_control = ruta_absoluta_relativa("Modelos_documentos/control_pagos_primera.docx")
            elif numero_armada == 'segunda':
                doc_control = ruta_absoluta_relativa("Modelos_documentos/control_pagos_segunda.docx")
            elif numero_armada == 'tercera':
                doc_control = ruta_absoluta_relativa("Modelos_documentos/control_pagos_tercera.docx")
            else:
                raise ValueError(f"Número de armada inválido: {numero_armada}")
            
            carpeta_final = os.path.join(carpeta_salida, "FASE FINAL", docente)
            os.makedirs(carpeta_final, exist_ok=True)

            nombre_archivo_control = f"CONTROL DE AVANCE - {docente} - {mes} {año}.docx"
            ruta_destino_control = os.path.join(carpeta_final, nombre_archivo_control)

            generar_control_avance(fila_control, doc_control, ruta_destino_control)

            print(f"{docente} - Control de pagos generado correctamente.")

        except Exception as e:
            print(f"Error procesando fila para {docente}: {e}")
            continue

