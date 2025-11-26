import pandas as pd
from PyPDF2 import PdfReader
import re
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from fuzzywuzzy import process

año_actual = datetime.datetime.now().year

# Configuración de email
EMAIL_CONFIG = {
    "remitente": "personalcontratado28.flch@unmsm.edu.pe",
    "password": "nbbr xttu qxqn tzej",
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 465
}

# Configuración alternativa (comentada)
# EMAIL_CONFIG = {
#     "remitente": "bolsistaceid01.flch@unmsm.edu.pe",
#     "password": "frsf imch edfs uwqy",
#     "smtp_server": "smtp.gmail.com",
#     "smtp_port": 465
# }

def extraer_nombre(pdf_path, debug=False):
    reader = PdfReader(pdf_path)
    texto = ""

    # Concatenar el texto de todas las páginas
    for pagina in reader.pages:
        texto += pagina.extract_text() + "\n"

    if debug:
        print(f"\n=== DEBUG: {os.path.basename(pdf_path)} ===")
        print(f"Texto extraído (primeros 500 caracteres):\n{texto[:500]}\n")

    # Función auxiliar para validar que el nombre tenga sentido
    def es_nombre_valido(nombre, min_len=3):
        if not nombre:
            return False
        partes = nombre.split(',')
        if len(partes) < 2:
            return False
        apellidos = partes[0].strip()
        nombres = partes[1].strip()
        # Contar solo letras (ignorar espacios)
        apellidos_letras = len([c for c in apellidos if c.isalpha()])
        nombres_letras = len([c for c in nombres if c.isalpha()])
        return apellidos_letras >= min_len and nombres_letras >= min_len

    candidatos = []

    # Patrón 1: Concepto:UNMSM seguido de números y nombre CON COMA
    match = re.search(r"Concepto:\s*UNMSM\s*\n?\s*\d+\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto, re.IGNORECASE)
    if match:
        nombre_raw = match.group(1)
        nombre_limpio = re.sub(r"\s+", " ", nombre_raw).strip()
        if es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 1 (Concepto:UNMSM con coma)", nombre_limpio))

    # Patrón 1b: Concepto:UNMSM seguido de números y nombre SIN COMA (formato: APELLIDOS NOMBRES)
    # Busca al menos 4 palabras en mayúsculas después del número
    match = re.search(r"Concepto:\s*UNMSM\s*\n?\s*(\d{8,})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto, re.IGNORECASE)
    if match:
        nombre_sin_coma = match.group(2).strip()
        # Filtrar palabras clave que no son nombres
        palabras_prohibidas = ["SERVICIO", "SERVICIOSERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS", 
                               "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "COORDINACION", "ATENCION",
                               "DIGITACION", "CLASIFICACION", "DESCRIPCION", "EVALUACION", "SOLICITUDES",
                               "RECEPCION", "ARCHIVO", "VISITAS", "DE", "LA", "Y", "EN"]
        
        primera_palabra = nombre_sin_coma.split()[0] if nombre_sin_coma.split() else ""
        
        # Validar que no comience con palabras prohibidas
        if primera_palabra not in palabras_prohibidas:
            # Convertir formato "APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2" a "APELLIDO1 APELLIDO2, NOMBRE1 NOMBRE2"
            palabras = nombre_sin_coma.split()
            if len(palabras) >= 3:
                # Tomar las primeras 2 palabras como apellidos y el resto como nombres
                apellidos = " ".join(palabras[:2])
                nombres = " ".join(palabras[2:])
                nombre_formateado = f"{apellidos}, {nombres}"
                nombre_limpio = re.sub(r"\s+", " ", nombre_formateado).strip()
                if es_nombre_valido(nombre_limpio, min_len=3):
                    candidatos.append(("Patrón 1b (Concepto:UNMSM sin coma)", nombre_limpio))

    # Patrón 2: números seguidos directamente de nombre CON COMA
    match = re.search(r"\d{8,}\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto)
    if match:
        nombre_raw = match.group(1)
        nombre_limpio = re.sub(r"\s+", " ", nombre_raw).strip()
        if es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 2 (números con coma)", nombre_limpio))

    # Patrón 2b: números largos seguidos de nombre SIN COMA
    match = re.search(r"\d{8,}\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto)
    if match:
        nombre_sin_coma = match.group(1).strip()
        # Filtrar palabras clave
        palabras_prohibidas = ["SERVICIO", "SERVICIOSERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS", 
                               "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "COORDINACION", "ATENCION",
                               "DIGITACION", "CLASIFICACION", "DESCRIPCION", "EVALUACION", "SOLICITUDES",
                               "RECEPCION", "ARCHIVO", "VISITAS", "DE", "LA", "Y", "EN", "MES", "DIA", "AÑO"]
        
        primera_palabra = nombre_sin_coma.split()[0] if nombre_sin_coma.split() else ""
        
        if primera_palabra not in palabras_prohibidas:
            palabras = nombre_sin_coma.split()
            if len(palabras) >= 3:
                apellidos = " ".join(palabras[:2])
                nombres = " ".join(palabras[2:])
                nombre_formateado = f"{apellidos}, {nombres}"
                nombre_limpio = re.sub(r"\s+", " ", nombre_formateado).strip()
                if es_nombre_valido(nombre_limpio, min_len=3):
                    candidatos.append(("Patrón 2b (números sin coma)", nombre_limpio))

    # Patrón 2c: Buscar específicamente después de RUC (11 dígitos) seguido del nombre
    match = re.search(r"\bRUC:.*?(\d{11})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto, re.IGNORECASE | re.DOTALL)
    if match:
        nombre_sin_coma = match.group(2).strip()
        palabras_prohibidas = ["SERVICIO", "SERVICIOSERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS", 
                               "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "COORDINACION", "ATENCION",
                               "DIGITACION", "CLASIFICACION", "DESCRIPCION", "EVALUACION", "SOLICITUDES",
                               "RECEPCION", "ARCHIVO", "VISITAS", "DE", "LA", "Y", "EN"]
        
        primera_palabra = nombre_sin_coma.split()[0] if nombre_sin_coma.split() else ""
        
        if primera_palabra not in palabras_prohibidas:
            palabras = nombre_sin_coma.split()
            if len(palabras) >= 3 and len(palabras) <= 6:  # Nombres típicos tienen 3-6 palabras
                apellidos = " ".join(palabras[:2])
                nombres = " ".join(palabras[2:])
                nombre_formateado = f"{apellidos}, {nombres}"
                nombre_limpio = re.sub(r"\s+", " ", nombre_formateado).strip()
                if es_nombre_valido(nombre_limpio, min_len=3):
                    candidatos.append(("Patrón 2c (después de RUC)", nombre_limpio))

    # Patrón 3: nombres con saltos de línea
    match = re.search(r"\b([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\s*,\s*\n?\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\b", texto)
    if match:
        apellidos = match.group(1).strip()
        nombres = match.group(2).strip()
        nombre_completo = f"{apellidos}, {nombres}"
        nombre_limpio = re.sub(r"\s+", " ", nombre_completo)
        if es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 3 (con saltos)", nombre_limpio))

    # Patrón 4: buscar todos los nombres con formato "APELLIDO(S), NOMBRE(S)" y elegir el más largo
    matches = re.findall(r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b", texto)
    for match in matches:
        nombre_limpio = re.sub(r"\s+", " ", match).strip()
        if es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 4 (fallback)", nombre_limpio))

    if debug and candidatos:
        print("Candidatos encontrados:")
        for patron, nombre in candidatos:
            print(f"  - {patron}: {nombre}")

    # Retornar el mejor candidato con prioridad por patrón y validación adicional
    if candidatos:
        # Filtrar candidatos que contengan palabras prohibidas
        palabras_prohibidas = ["SERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS", 
                               "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "OPERADOR",
                               "COORDINACION", "ATENCION", "ORDEN", "DE LA"]
        
        candidatos_validos = []
        for patron, nombre in candidatos:
            # Verificar que el nombre no contenga palabras prohibidas
            nombre_upper = nombre.upper()
            tiene_prohibida = any(palabra in nombre_upper for palabra in palabras_prohibidas)
            
            # Verificar longitud razonable (nombres típicos tienen entre 15 y 60 caracteres)
            longitud_ok = 15 <= len(nombre) <= 60
            
            if not tiene_prohibida and longitud_ok:
                candidatos_validos.append((patron, nombre))
        
        # Si no quedan candidatos válidos después del filtro, usar todos
        if not candidatos_validos:
            candidatos_validos = candidatos
        
        # Priorizar por tipo de patrón (los primeros son más confiables)
        prioridad = {
            "Patrón 1 (Concepto:UNMSM con coma)": 1,
            "Patrón 2 (números con coma)": 2,
            "Patrón 1b (Concepto:UNMSM sin coma)": 3,
            "Patrón 2b (números sin coma)": 4,
            "Patrón 2c (después de RUC)": 5,
            "Patrón 3 (con saltos)": 6,
            "Patrón 4 (fallback)": 7
        }
        
        # Ordenar por prioridad y luego por longitud
        mejor = min(candidatos_validos, key=lambda x: (prioridad.get(x[0], 99), -len(x[1])))
        
        if debug:
            print(f"Seleccionado: {mejor[1]}\n")
        return mejor[1]

    if debug:
        print("No se encontró ningún nombre válido\n")
    
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



def generar_firma_html() -> str:
    return """
    <p>Atentamente,</p>

    <p style="font-size: 13px; font-weight: bold; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">C.P.C. María Rivera Vidal</p>
    <p style="font-size: 10px; font-style: sans-serif; color:rgb(82,82,82); margin:0cm 0cm 0.0001pt; line-height: normal">Responsable de la Coordinación de Procesos Administrativos </p>

    <span style="font-size:10px; font-weight: bold; font-style: sans-serif; color:rgb(11,83,148); margin-bottom: 0cm; line-height: normal">Centro de Idiomas de la Universidad Nacional Mayor de San Marcos</span>

    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Contacto: (01) 619 7000 Anexo 2848</p>
    <p style="font-size:10px; font-style: sans-serif; color:rgb(51,51,51); margin-bottom: 0cm; line-height: normal">Av. Universitaria,  Calle Germán Amézaga N.° 375. Ciudad Universitaria, Lima.</p>
"""

def generar_cuerpo_correo_docente_html(mes: str, anio: str, servicio: str) -> str:
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

        {generar_firma_html()}
    </body>
</html>
"""

def generar_cuerpo_correo_administrativo_html(mes: str, anio: str) -> str:
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

        {generar_firma_html()}
    </body>
</html>
"""

def adjuntar_pdf(msg: MIMEMultipart, pdf_path: str) -> None:
    with open(pdf_path, 'rb') as adjunto:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(adjunto.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
        msg.attach(parte)

def enviar_correo(destinatario: str, asunto: str, cuerpo_html: str, pdf_path: str, nombre: str) -> None:
    # Crear mensaje
    msg = MIMEMultipart()
    msg['From'] = EMAIL_CONFIG["remitente"]
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo_html, 'html'))

    # Adjuntar PDF
    adjuntar_pdf(msg, pdf_path)

    # Enviar
    with smtplib.SMTP_SSL(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"]) as servidor:
        servidor.login(EMAIL_CONFIG["remitente"], EMAIL_CONFIG["password"])
        servidor.sendmail(EMAIL_CONFIG["remitente"], [destinatario], msg.as_string())

    print(f"Correo enviado a {nombre}.")

def enviar_correo_docente(nombre: str, pdf_path: str, destinatario: str, mes: str, servicio: str) -> None:
    nombre_formato = nombre.split(",")[0].strip()
    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {año_actual} - {nombre_formato}"
    cuerpo_html = generar_cuerpo_correo_docente_html(mes, año_actual, servicio)
    
    enviar_correo(destinatario, asunto, cuerpo_html, pdf_path, nombre)

def enviar_correo_administrativo(nombre: str, pdf_path: str, destinatario: str, mes: str) -> None:
    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {año_actual}"
    cuerpo_html = generar_cuerpo_correo_administrativo_html(mes, año_actual)
    
    enviar_correo(destinatario, asunto, cuerpo_html, pdf_path, nombre)

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
            print(f"⚠ No se encontró nombre válido en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_docente in mapeo_correos:
            datos = mapeo_correos[nombre_docente]
            if datos["servicio"]:
                resultados.append(datos)
                print(f"{datos['nombre']} - {datos['correo']} - {datos['servicio']}")
            else:
                print(f"⚠ No se encontró servicio para {datos['nombre']}.")
        else:
            print(f"⚠ Coincidencia baja para '{nombre_docente}' (archivo: {os.path.basename(pdf_path)}), omitido.")

    return resultados

def procesar_correos_administrativos(ruta_excel, hoja, lista_pdfs, debug=False):
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    # Asegurar que la columna 'Docente' exista
    if 'Docente' not in df.columns or 'Correo Institucional' not in df.columns:
        raise ValueError("El Excel debe contener columnas 'Docente' y 'Correo Institucional'.")

    nombres_excel = df['Docente'].astype(str).tolist()

    # Crear mapeo con debug
    mapeo_correos = {}
    for pdf_path in lista_pdfs:
        nombre_extraido = extraer_nombre(pdf_path, debug=debug)
        if not nombre_extraido:
            continue
            
        if nombre_extraido not in mapeo_correos:
            resultado = process.extractOne(nombre_extraido, nombres_excel)
            if resultado and resultado[1] >= 80:
                mejor_match = resultado[0]
                fila = df[df['Docente'] == mejor_match]
                if not fila.empty:
                    correo = fila['Correo Institucional'].values[0]
                    mapeo_correos[nombre_extraido] = {
                        "pdf_path": pdf_path,
                        "nombre": mejor_match,
                        "correo": correo
                    }

    resultados = []
    for pdf_path in lista_pdfs:
        nombre_administrativo = extraer_nombre(pdf_path, debug=debug)
        if not nombre_administrativo:
            print(f"⚠ No se encontró nombre válido en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_administrativo in mapeo_correos:
            datos = mapeo_correos[nombre_administrativo]
            resultados.append(datos)
            print(f"{datos['nombre']} - {datos['correo']}")
        else:
            print(f"⚠ Coincidencia baja para '{nombre_administrativo}' (archivo: {os.path.basename(pdf_path)}), omitido.")

    return resultados