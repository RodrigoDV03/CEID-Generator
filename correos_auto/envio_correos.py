import pandas as pd
from PyPDF2 import PdfReader
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from fuzzywuzzy import process

def extraer_nombre(pdf_path):
    reader = PdfReader(pdf_path)
    texto = ""

    # Concatenar el texto de todas las páginas
    for pagina in reader.pages:
        texto += pagina.extract_text() + "\n"

    # Buscar patrón específico: Concepto:UNMSM seguido de números y nombre
    # Patrón mejorado que busca "Concepto:" seguido opcionalmente de "UNMSM", números y nombre
    match = re.search(r"Concepto:\s*UNMSM\s*\n?\s*\d+\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto, re.IGNORECASE)
    if match:
        nombre_raw = match.group(1)
        # Limpiar espacios múltiples
        return re.sub(r"\s+", " ", nombre_raw).strip()

    # Patrón alternativo: buscar números seguidos directamente de nombre (sin "Concepto:")
    match = re.search(r"\d{8,}\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto)
    if match:
        nombre_raw = match.group(1)
        # Limpiar espacios múltiples
        return re.sub(r"\s+", " ", nombre_raw).strip()

    # Patrón original como fallback
    match = re.search(r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b", texto)
    if match:
        nombre_raw = match.group(1)
        # Limpiar espacios múltiples
        return re.sub(r"\s+", " ", nombre_raw).strip()

    return None

def extraer_servicios(pdf_path):

    reader = PdfReader(pdf_path)
    texto = ""

    # Concatenar el texto de todas las páginas
    for pagina in reader.pages:
        texto += pagina.extract_text() + "\n"

    lineas = texto.splitlines()

    # Buscar índice de inicio de la tabla
    idx_inicio = None
    for i, linea in enumerate(lineas):
        if "Código Unid. Med." in linea and "Descripción" in linea:
            idx_inicio = i
            break

    if idx_inicio is None:
        return None

    # Patrón mejorado: permite "28 horas" o "28horas"
    patron = re.compile(r"^\d{1,2}\s*horas\s+de\s+.*", re.IGNORECASE)

    horas = []
    for linea in lineas[idx_inicio:]:
        if patron.match(linea.strip()):
            horas.append(re.sub(r"\s+", " ", linea.strip()))  # limpiar espacios extra
        elif horas:  # si ya empezó y cortó
            break

    if not horas:
        return None

    # Formatear con comas y "y"
    if len(horas) > 1:
        return ", ".join(horas[:-1]) + " y " + horas[-1]
    else:
        return horas[0]



def generar_cuerpo_correo_docente_html(mes, anio, servicio):
    return f"""
<html>
  <body style="font-family: Verdana; font-size: 10pt;">
    <p>Buen día:</p>

    <p>
      Adjunto su orden de servicio correspondiente al mes de {mes} {anio}. Con este documento, ya puede proceder con la emisión de su recibo por honorarios. Para evitar retrasos en el pago, tenga en cuenta lo siguiente:
    </p>

    <ul>
      <li>
        El concepto del recibo por honorarios es: <span style="font-weight: bold; color: #073763;"> Servicio de dictado de {servicio}, BAJO LA MODALIDAD HÍBRIDA.</span>
      </li>

      <li>
        El pago se realizará a crédito, con un plazo de vencimiento de <span style="font-weight: bold; color: #073763;">40 días</span> desde la fecha de emisión del recibo en la plataforma SUNAT.
      </li>
    </ul>

    <p>
      Una vez emitido, envíe el recibo en <span style="font-weight: bold; color: #073763;">formato PDF</span> como respuesta a este mismo correo, sin generar un nuevo hilo.
    </p>

    <p>Atentamente,</p>

    <p style="font-size: 13px; font-weight: bold; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">C.P.C. María Rivera Vidal</p>
    <p style="font-size: 10px; font-weight: bold; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">Responsable de la Coordinación de Procesos Administrativos </p>

    <span style="font-size:10px; font-weight: bold; font-style: sans-serif; color:rgb(11,83,148); margin-bottom: 0cm; line-height: normal">Centro de Idiomas de la Universidad Nacional Mayor de San Marcos</span>

    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Contacto: (01) 619 7000 Anexo 2848</p>
    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Av. Universitaria,  Calle Germán Amézaga N.° 375. Ciudad Universitaria, Lima.</p>
  </body>
</html>
"""

def generar_cuerpo_correo_administrativo_html(mes, anio):
    return f"""
<html>
  <body style="font-family: Verdana; font-size: 10pt;">
    <p>Buen día:</p>

    <p>
      Adjunto su orden de servicio correspondiente al mes de {mes} {anio}. Con este documento, ya puede proceder con la emisión de su recibo por honorarios.
    </p>

    <p>
      Una vez emitido, envíe el recibo en <span style="font-weight: bold; color: #073763;">formato PDF</span> como respuesta a este mismo correo, sin generar un nuevo hilo.
    </p>

    <p>Atentamente,</p>

    <p style="font-size: 13px; font-weight: bold; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">C.P.C. María Rivera Vidal</p>
    <p style="font-size: 10px; font-weight: bold; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">Responsable de la Coordinación de Procesos Administrativos </p>

    <span style="font-size:10px; font-weight: bold; font-style: sans-serif; color:rgb(11,83,148); margin-bottom: 0cm; line-height: normal">Centro de Idiomas de la Universidad Nacional Mayor de San Marcos</span>

    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Contacto: (01) 619 7000 Anexo 2848</p>
    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Av. Universitaria,  Calle Germán Amézaga N.° 375. Ciudad Universitaria, Lima.</p>
  </body>
</html>
"""

def enviar_correo_docente(nombre, pdf_path, destinatario, mes, anio, servicio):
    remitente = "personalcontratado28.flch@unmsm.edu.pe"
    password = "vfrl usic kmfm fyah"

    # remitente = "bolsistaceid01.flch@unmsm.edu.pe"
    # password = "frsf imch edfs uwqy"

    nombre_formato = nombre.split(",")[0].strip()

    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {anio} - {nombre_formato}"

    cuerpo_html = generar_cuerpo_correo_docente_html(mes, anio, servicio)

    # Crear mensaje
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo_html, 'html'))

    # Adjuntar PDF
    with open(pdf_path, 'rb') as adjunto:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(adjunto.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
        msg.attach(parte)

    # Enviar
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
        servidor.login(remitente, password)
        servidor.sendmail(remitente, [destinatario], msg.as_string())

    print(f"Correo enviado a {nombre}.")

def enviar_correo_administrativo(nombre, pdf_path, destinatario, mes, anio):
    remitente = "personalcontratado28.flch@unmsm.edu.pe"
    password = "vfrl usic kmfm fyah"

    # remitente = "bolsistaceid01.flch@unmsm.edu.pe"
    # password = "frsf imch edfs uwqy"

    # nombre_formato = nombre.split(",")[0].strip()

    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {anio}"

    cuerpo_html = generar_cuerpo_correo_administrativo_html(mes, anio)

    # Crear mensaje
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo_html, 'html'))

    # Adjuntar PDF
    with open(pdf_path, 'rb') as adjunto:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(adjunto.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
        msg.attach(parte)

    # Enviar
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
        servidor.login(remitente, password)
        servidor.sendmail(remitente, [destinatario], msg.as_string())

    print(f"Correo enviado a {nombre}.")

def crear_mapeo_correos(lista_pdfs, nombres_excel, df, tipo_correo="docente"):
    mapeo = {}
    
    for pdf_path in lista_pdfs:
        nombre_extraido = extraer_nombre(pdf_path)
        if not nombre_extraido:
            continue
            
        if nombre_extraido not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre_extraido, nombres_excel)
            if resultado and resultado[1] >= 80:
                mejor_match = resultado[0]
                fila = df[df['Docente'] == mejor_match]
                if not fila.empty:
                    correo = fila['Correo Institucional'].values[0]
                    if tipo_correo == "docente":
                        servicio = extraer_servicios(pdf_path)
                        mapeo[nombre_extraido] = {
                            "pdf_path": pdf_path,
                            "nombre": mejor_match,
                            "correo": correo,
                            "servicio": servicio
                        }
                    else:
                        mapeo[nombre_extraido] = {
                            "pdf_path": pdf_path,
                            "nombre": mejor_match,
                            "correo": correo
                        }
    
    return mapeo

def procesar_correos_docente(ruta_excel, hoja, lista_pdfs):
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    # Asegurar que la columna 'Docente' exista
    if 'Docente' not in df.columns or 'Correo Institucional' not in df.columns:
        raise ValueError("El Excel debe contener columnas 'Docente' y 'Correo Institucional'.")

    nombres_excel = df['Docente'].astype(str).tolist()

    # Crear mapeo fuzzy una sola vez
    mapeo_correos = crear_mapeo_correos(lista_pdfs, nombres_excel, df, "docente")

    resultados = []
    for pdf_path in lista_pdfs:
        nombre_docente = extraer_nombre(pdf_path)
        if not nombre_docente:
            print(f"⚠ No se encontró nombre en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_docente in mapeo_correos:
            datos = mapeo_correos[nombre_docente]
            if datos["servicio"]:
                resultados.append(datos)
                print(f"{datos['nombre']} - {datos['correo']} - {datos['servicio']}")
            else:
                print(f"⚠ No se encontró servicio para {datos['nombre']}.")
        else:
            print(f"⚠ Coincidencia baja para '{nombre_docente}', omitido.")

    return resultados

def procesar_correos_administrativos(ruta_excel, hoja, lista_pdfs):
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    # Asegurar que la columna 'Docente' exista
    if 'Docente' not in df.columns or 'Correo Institucional' not in df.columns:
        raise ValueError("El Excel debe contener columnas 'Docente' y 'Correo Institucional'.")

    nombres_excel = df['Docente'].astype(str).tolist()

    mapeo_correos = crear_mapeo_correos(lista_pdfs, nombres_excel, df, "administrativo")

    resultados = []
    for pdf_path in lista_pdfs:
        nombre_administrativo = extraer_nombre(pdf_path)
        if not nombre_administrativo:
            print(f"⚠ No se encontró nombre en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_administrativo in mapeo_correos:
            datos = mapeo_correos[nombre_administrativo]
            resultados.append(datos)
            print(f"{datos['nombre']} - {datos['correo']}")
        else:
            print(f"⚠ Coincidencia baja para '{nombre_administrativo}', omitido.")

    return resultados