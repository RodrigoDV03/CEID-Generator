import os
import pandas as pd
import datetime
from fases.functions import *
from docx2pdf import convert

def procesar_planilla_fase_inicial(planilla_path, hoja_seleccionada, carpeta_destino, mes, numero_armada, tipo_fase_inicial):
    año_actual = datetime.datetime.now().year

    datos = pd.read_excel(planilla_path, sheet_name=hoja_seleccionada)
    datos.columns = datos.columns.str.strip()

    carpeta_principal = os.path.join(carpeta_destino, 'FASE INICIAL')
    os.makedirs(carpeta_principal, exist_ok=True)

    for i, fila in enumerate(datos.itertuples(index=False), start=1):
        docente = str(getattr(fila, "Docente", "N/A"))

        if docente == "N/A":
            print(f"Fila {i}: No se pudo obtener el nombre del docente. Columnas disponibles: {list(datos.columns)}")
            continue
        
        nombre_docente = limpiar_nombre_archivo(docente)
        carpeta_docente = os.path.join(carpeta_principal, nombre_docente)
        os.makedirs(carpeta_docente, exist_ok=True)

        categoria_valor = getattr(fila, "Categoria_monto", 1)
        if pd.isna(categoria_valor):
            categoria_valor = 1
        else:
            categoria_valor = float(categoria_valor)
        ruc = limpiar_numero(getattr(fila, "N_Ruc", ""))
        if tipo_fase_inicial == "administrativo":
            descripcion = str(getattr(fila, "Curso", ""))
            finalidad_publica = str(getattr(fila, "Finalidad_publica", ""))
            formacion_academica = str(getattr(fila, "Formacion_academica", ""))
            experiencia_laboral = str(getattr(fila, "Experiencia_laboral", ""))
            requisitos_adi = str(getattr(fila, "Requisitos_adicional", ""))
            actividades_admin = str(getattr(fila, "Actividades_admin", ""))
        else:
            descripcion_raw = str(getattr(fila, "Curso", ""))
            descripcion = redactar_cursos(descripcion_raw)
        disenio_examenes = float(getattr(fila, "Disenio_examenes", 0))
        disenio_cant_horas = disenio_examenes / categoria_valor
        horas_disenio = f"{int(round(disenio_cant_horas))} horas de diseño de exámenes"
        clasif_valor = int(getattr(fila, "Examen_clasif", 0))
        direccion = str(getattr(fila, "Domicilio_docente", '')).strip()
        correo = str(getattr(fila, "Correo_personal", ''))
        celular = limpiar_numero(getattr(fila, "Numero_celular", ""))
        dni_docente = limpiar_numero(getattr(fila, "Numero_dni", ""))
        if len(dni_docente) < 8:
            dni_docente = dni_docente.zfill(8)



        clasif_cant_horas = clasif_valor / categoria_valor
        if clasif_cant_horas == 1:
            horas_clasif = f"{int(round(clasif_cant_horas))} hora de examen de clasificación"
        else:
            horas_clasif = f"{int(round(clasif_cant_horas))} horas de examen de clasificación"

        if tipo_fase_inicial == "administrativo":
            descripcion_final = f"{descripcion}"
        else:   
            if clasif_valor == 0:
                descripcion_final = f"{descripcion} y {horas_disenio}"
            else:
                descripcion_final = f"{descripcion}, {horas_disenio} y {horas_clasif}"

        monto_categoria_letras = monto_a_letras(categoria_valor)
        monto_total = getattr(fila, "Total_pago", 0)
        monto_total_letras = monto_a_letras(monto_total)
        tipo_contrato = getattr(fila, "Estado_docente", "N/A")
        nro_contrato_val = getattr(fila, "Nro_Contrato", "N/A")
        try:
            nro_contrato = str(int(float(nro_contrato_val)))
        except (ValueError, TypeError):
            nro_contrato = str(nro_contrato_val)

        if tipo_fase_inicial == "administrativo":
            modalidad_servicio = "presencial"
        else:
            modalidad_servicio = "híbrida"

        # -------- GENERAR OFICIO --------
        if tipo_contrato == "CONTRATO":
            ruta_oficio = ruta_absoluta_relativa('./Modelos_documentos/oficio_contrato.docx')
        elif tipo_contrato == "TERCERO":
            ruta_oficio = ruta_absoluta_relativa('./Modelos_documentos/oficio_tercero.docx')
        else:
            ruta_oficio = None

        reemplazos_oficio = {
            "Nro_Contrato": nro_contrato,
            "docente": docente,
            "descripcion": descripcion_final,
            "categoria": f"S/. {categoria_valor:,.2f} ({monto_categoria_letras})",
            "monto_subtotal": f"S/. {monto_total:,.2f} ({monto_total_letras})",
            "numero_armada": numero_armada,
            "modalidad_servicio": modalidad_servicio
        }
        ruta_salida_oficio = os.path.join(carpeta_docente, f"OFICIO - {nombre_docente} - {mes} {año_actual}.docx")
        generar_documento(ruta_oficio, reemplazos_oficio, ruta_salida_oficio)
        if tipo_fase_inicial == "administrativo":
            if os.path.exists(ruta_salida_oficio):
                doc = Document(ruta_salida_oficio)
                for parrafo in doc.paragraphs:
                    for run in parrafo.runs:
                        if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                            run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                for tabla in doc.tables:
                    for fila in tabla.rows:
                        for celda in fila.cells:
                            for parrafo in celda.paragraphs:
                                for run in parrafo.runs:
                                    if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                        run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                doc.save(ruta_salida_oficio)

        # -------- GENERAR TDR --------
        if tipo_contrato == "TERCERO":
            tipo_tdr = str(getattr(fila, "Categoria_letra", "")).strip().upper()
            ruta_tdr = ruta_absoluta_relativa(f'./Modelos_documentos/tdr_tipo{tipo_tdr}_.docx')

            reemplazos_tdr = {
                "descripcion": descripcion_final,
                "categoria": f"S/. {categoria_valor:,.2f} ({monto_categoria_letras})",
                "monto_subtotal": f"S/. {monto_total:,.2f} ({monto_total_letras})",
                "modalidad_servicio": modalidad_servicio,
            }

            if tipo_fase_inicial == "administrativo":
                ruta_tdr = ruta_absoluta_relativa('./Modelos_documentos/tdr_administrativo.docx')

                reemplazos_tdr = {
                    "descripcion": descripcion_final,
                    "finalidad_publica": finalidad_publica,
                    "formacion_academica": formacion_academica,
                    "experiencia_laboral": experiencia_laboral,
                    "requisitos_adicional": requisitos_adi,
                    "actividades_admin": actividades_admin,
                    "categoria": f"S/. {categoria_valor:,.2f} ({monto_categoria_letras})",
                    "monto_subtotal": f"S/. {monto_total:,.2f} ({monto_total_letras})",
                    "modalidad_servicio": modalidad_servicio
                }

            ruta_salida_tdr = os.path.join(carpeta_docente, f"TDR - {nombre_docente} - {mes} {año_actual}.docx")
            generar_documento(ruta_tdr, reemplazos_tdr, ruta_salida_tdr)
            if tipo_fase_inicial == "administrativo":
                if os.path.exists(ruta_salida_tdr):
                    doc = Document(ruta_salida_tdr)
                    for parrafo in doc.paragraphs:
                        for run in parrafo.runs:
                            if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                    for tabla in doc.tables:
                        for fila in tabla.rows:
                            for celda in fila.cells:
                                for parrafo in celda.paragraphs:
                                    for run in parrafo.runs:
                                        if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                            run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                    doc.save(ruta_salida_tdr)
            ruta_salida_tdr_pdf = ruta_salida_tdr.replace('.docx', '.pdf')
            convert(ruta_salida_tdr, ruta_salida_tdr_pdf)

        # -------- GENERAR COTIZACIÓN --------

        if tipo_fase_inicial == "administrativo" or (tipo_fase_inicial != "administrativo" and tipo_contrato == "TERCERO"):
            ruta_cotizacion = ruta_absoluta_relativa('./Modelos_documentos/modelo_cotizacion.docx')

            if tipo_fase_inicial != "administrativo":
                actividades_admin = """-	Dictar clases
    -	Preparar las clases
    -	Evaluar a los alumnos
    -	Diseñar exámenes
    -	Entregar acta de nota
    -	Presentar informe de dictado de curso
    """
                ruta_firma = ruta_absoluta_relativa(f'firmas_docentes/{nombre_docente}.png')
            else:
                ruta_firma = ruta_absoluta_relativa(f'firmas_admin/{nombre_docente}.png')


            reemplazos = {
                "nombre_docente": docente,
                "direccion_cot": f"Dirección: {direccion}",
                "ruc_docente_cot": f"RUC N.º {ruc}",
                "correo_docente_cot": f"Correo: {correo}",
                "celular_cot": f"Teléfono: {celular}",
                "descripcion_servicio": descripcion_final,
                "actividades_admin": actividades_admin,
                "categoria_monto": f"S/. {categoria_valor:,.2f} ({monto_categoria_letras})",
                "monto_subtotal": f"S/. {monto_total:,.2f} ({monto_total_letras})",
                "dni_cot": f"DNI: {dni_docente}",
                "modalidad_servicio": modalidad_servicio,
            }

            ruta_salida_cot = os.path.join(carpeta_docente, f"COTIZACIÓN - {nombre_docente} - {mes} {año_actual}.docx")
            generar_documento(ruta_cotizacion, reemplazos, ruta_salida_cot, ruta_firma)
            if tipo_fase_inicial == "administrativo":
                if os.path.exists(ruta_salida_cot):
                    doc = Document(ruta_salida_cot)
                    for parrafo in doc.paragraphs:
                        for run in parrafo.runs:
                            if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                    for tabla in doc.tables:
                        for fila in tabla.rows:
                            for celda in fila.cells:
                                for parrafo in celda.paragraphs:
                                    for run in parrafo.runs:
                                        if ", monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text or "Monto por hora: S/. 1.00 (uno y 00/100 soles)" in run.text:
                                            run.text = run.text.replace(", monto por hora: S/. 1.00 (uno y 00/100 soles)", "").replace("Monto por hora: S/. 1.00 (uno y 00/100 soles)", "")
                    doc.save(ruta_salida_cot)

        print(f"{docente} - Documentos generados correctamente.")