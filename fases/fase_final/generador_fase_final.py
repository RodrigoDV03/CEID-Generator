import os
import pandas as pd
from fases.functions import *

def generador_conformidad(fila, ruta_conformidad, ruta_destino, numero_armada, tipo_fase_final):

    docente = str(getattr(fila, "Docente", "N/A"))
    nombre_docente = docente

    if not os.path.exists(ruta_conformidad):
        raise FileNotFoundError(f"No se encontró la plantilla: {ruta_conformidad}")

    ruc = limpiar_numero(getattr(fila, "N_Ruc", ""))
    descripcion_raw = str(getattr(fila, "Curso", ""))
    if tipo_fase_final == "administrativo":
        descripcion = str(getattr(fila, "Curso", "N/A"))
    else:
        descripcion = redactar_cursos(descripcion_raw)
    cant_cursos = int(getattr(fila, "Cantidad_cursos", 0))
    disenio_cant_horas = cant_cursos * 4
    horas_disenio = f"{int(round(disenio_cant_horas))} horas de diseño de exámenes"
    clasif_valor = int(getattr(fila, "Examen_clasif", 0))

    categoria_valor = getattr(fila, "Categoria_monto", 1)
    if pd.isna(categoria_valor):
        categoria_valor = 1
    else:
        categoria_valor = float(categoria_valor)

    clasif_cant_horas = clasif_valor / categoria_valor
    if clasif_cant_horas == 1:
        horas_clasif = f"{int(round(clasif_cant_horas))} hora de examen de clasificación"
    else:
        horas_clasif = f"{int(round(clasif_cant_horas))} horas de examen de clasificación"


    if tipo_fase_final == "administrativo":
        descripcion_final = descripcion
    else:
        if clasif_valor == 0:
            descripcion_final = f"{descripcion} y {horas_disenio}"
        else:
            descripcion_final = f"{descripcion}, {horas_disenio} y {horas_clasif}"

    monto_categoria_letras = monto_a_letras(categoria_valor)
    monto_total = getattr(fila, "Subtotal_pago", 0)
    monto_total_letras = monto_a_letras(monto_total)
    nro_contrato_val = getattr(fila, "Nro_Contrato", "")
    
    try:
        nro_contrato = str(int(float(nro_contrato_val)))
    except (ValueError, TypeError):
        nro_contrato = str(nro_contrato_val)

    if tipo_fase_final == "administrativo":
        modalidad_servicio = "presencial"
    else:
        modalidad_servicio = "híbrida"

    reemplazos_conformidad = {
        "nombre": str(nombre_docente),
        "ruc": str(ruc),
        "descripcion_cursos": str(descripcion_final),
        "monto_subtotal": f"S/. {monto_total:,.2f} ({str(monto_total_letras)})",
        "monto_hora": f"S/. {categoria_valor:,.2f} ({str(monto_categoria_letras)})",
        "Nro_Contrato": str(nro_contrato),
        "numero_armada": str(numero_armada),
        "modalidad_servicio": str(modalidad_servicio)
    }

    carpeta_final = os.path.dirname(ruta_destino)
    os.makedirs(carpeta_final, exist_ok=True)
    generar_documento(ruta_conformidad, reemplazos_conformidad, ruta_destino)
    if tipo_fase_final == "administrativo":
        if os.path.exists(ruta_destino):
            doc = Document(ruta_destino)
            for parrafo in doc.paragraphs:
                for run in parrafo.runs:
                    if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                        run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        for parrafo in celda.paragraphs:
                            for run in parrafo.runs:
                                if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                    run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
            doc.save(ruta_destino)
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

    reemplazos_armada = {
        "Nombre_Docente": str(nombre_docente),
        "Nro_Contrato": numero_contrato_str,
        "Idioma_Docente": str(idioma_docente),
        "Monto_Total": f"S/. {monto_total:,.2f}",
        "Total_Primera": f"S/. {total_primera:,.2f}",
        "Total_Segunda": f"S/. {total_segunda:,.2f}",
        "Total_Tercera": f"S/. {total_tercera:,.2f}",
        "Saldo_Restante": f"S/. {saldo_restante:,.2f}",
        "Saldo_Primera": f"S/. {saldo_primera:,.2f}",
        "Saldo_Segunda": f"S/. {saldo_segunda:,.2f}",
    }

    carpeta_final = os.path.dirname(ruta_destino)
    os.makedirs(carpeta_final, exist_ok=True)
    generar_documento(doc_control, reemplazos_armada, ruta_destino)
    return ruta_destino


def procesar_planilla_fase_final(planilla_path, excel_control_pagos, hoja, carpeta_salida, mes, año, numero_armada, tipo_fase_final):
    df = pd.read_excel(planilla_path, sheet_name=hoja)

    if tipo_fase_final == "planilla docente (con contrato)":
        df_control = pd.read_excel(excel_control_pagos, sheet_name=0, header=1)

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
            nombre_archivo_conformidad = f"CONFORMIDAD - {docente} - {mes} {año}.docx"
            ruta_destino_conformidad = os.path.join(carpeta_final, nombre_archivo_conformidad)
            generador_conformidad(fila, ruta_conformidad, ruta_destino_conformidad, numero_armada, tipo_fase_final)

            print(f"{docente} - Documento de conformidad generado correctamente.")
        except Exception as e:
            print(f"Error procesando fila para {docente}: {e}")
            continue


    if tipo_fase_final == "planilla docente (con contrato)":
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

