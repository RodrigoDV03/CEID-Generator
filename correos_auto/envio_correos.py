import pandas as pd
import pdfplumber
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from fuzzywuzzy import process

def extraer_nombre_docente(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto = ""
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"

    # Buscar patrón de nombre en el PDF
    match = re.search(r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b", texto)
    if match:
        nombre_raw = match.group(1)
        # Limpiar espacios múltiples
        return re.sub(r"\s+", " ", nombre_raw).strip()
    return None

def extraer_servicios(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:
        texto = ""
        for pagina in pdf.pages:
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



def generar_cuerpo_correo_html(mes, anio, servicio):
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

    <p style="font-weight: bold; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">Rodrigo Estéfano, Dávila Vásquez</p>
    <p style="font-weight: bold; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">Coordinación de procesos administrativos</p>

    <span style="font-size:7.5pt; font-weight: bold;color:rgb(11,83,148); margin-bottom: 0cm; line-height: normal">Centro de Idiomas de la Universidad Nacional Mayor de San Marcos</span>

    <p style="font-size:7.5pt;color:rgb(68,68,68); margin-bottom: 0cm; line-height: normal">Contacto: (01) 619 7000 Anexo 2848</p>
    <p style="font-size:7pt;color:rgb(68,68,68); margin-bottom: 0cm; line-height: normal">Av. Universitaria,  Calle Germán Amézaga N.° 375. Ciudad Universitaria, Lima.</p>
  </body>
</html>
"""


def enviar_correo(pdf_path, destinatario, mes, anio, servicio):
    # remitente = "personalcontratado28.flch@unmsm.edu.pe"
    # cc = "coordinacionsistemasceid.flch@unmsm.edu.pe"
    # password = "vfrl usic kmfm fyah"


    remitente = "bolsistaceid01.flch@unmsm.edu.pe"
    cc = "rodrodv03@gmail.com"
    password = "frsf imch edfs uwqy"

    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {anio}"
    cuerpo_html = generar_cuerpo_correo_html(mes, anio, servicio)

    # Crear mensaje
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Cc'] = cc
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
        servidor.sendmail(remitente, [destinatario, cc], msg.as_string())

    print(f"Correo enviado a {destinatario} con CC a {cc}")

def procesar_correos(ruta_excel, hoja, mes, anio, lista_pdfs):
    df = pd.read_excel(ruta_excel, sheet_name=hoja)

    nombres_excel = df['Docente'].astype(str).tolist()

    for pdf_path in lista_pdfs:
        nombre_docente = extraer_nombre_docente(pdf_path)
        if not nombre_docente:
            print(f"No se encontró nombre en {pdf_path}, omitido.")
            continue

        # Emparejamiento difuso con precisión mínima de 90
        mejor_match, score = process.extractOne(nombre_docente, nombres_excel)
        if score < 90:
            print(f"No se encontró coincidencia suficiente para '{nombre_docente}' (score={score}), omitido.")
            continue

        docente_row = df[df['Docente'] == mejor_match]
        correo_destinatario = docente_row['Correo Institucional'].values[0]
        servicio = extraer_servicios(pdf_path)
        enviar_correo(pdf_path, correo_destinatario, mes, anio, servicio)
        print(f"Profesor: {mejor_match}")
        print(f"Correo: {correo_destinatario}")
        print(f"Servicio: {servicio}")
        print(f"-------------------------------------------------------------")
        servicio = None
