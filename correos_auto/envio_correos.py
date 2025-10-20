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
from utils.config_manager import get_email_config
from repositories import excel_repo

# Configuración global
CURRENT_YEAR = datetime.datetime.now().year
SIMILARITY_THRESHOLD = 80  # Umbral de similitud para fuzzy matching

def get_email_settings():
    email_config = get_email_config()
    return {
        "remitente": email_config.get("remitente", "personalcontratado28.flch@unmsm.edu.pe"),
        "password": email_config.get("password", "nbbr xttu qxqn tzej"),
        "smtp_server": email_config.get("smtp_server", "smtp.gmail.com"),
        "smtp_port": email_config.get("smtp_port", 465)
    }

# Configuración de email singleton
EMAIL_CONFIG = get_email_settings()

def extraer_nombre(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        texto = ""

        # Concatenar texto de todas las páginas
        for pagina in reader.pages:
            texto += pagina.extract_text() + "\n"

        # Patrones de búsqueda ordenados por precisión
        patrones = [
            # Patrón específico: Concepto:UNMSM seguido de números y nombre
            r"Concepto:\s*UNMSM\s*\n?\s*\d+\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)",
            # Patrón alternativo: números seguidos directamente de nombre
            r"\d{8,}\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)",
            # Patrón genérico como fallback
            r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b"
        ]

        for patron in patrones:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                nombre_raw = match.group(1)
                # Limpiar espacios múltiples y normalizar
                return re.sub(r"\s+", " ", nombre_raw).strip().upper()

        return None
    
    except Exception as e:
        print(f"⚠️ Error extrayendo nombre de {pdf_path}: {e}")
        return None

def extraer_servicios(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        texto = ""

        # Concatenar texto de todas las páginas
        for pagina in reader.pages:
            texto += pagina.extract_text() + "\n"

        lineas = texto.splitlines()

        # Buscar tabla de servicios
        idx_inicio = _find_services_table_start(lineas)
        if idx_inicio is None:
            return None

        # Extraer servicios de enseñanza
        servicios = _extract_teaching_services(lineas[idx_inicio:])
        
        if not servicios:
            return None

        # Formatear lista de servicios
        return _format_services_list(servicios)
    
    except Exception as e:
        print(f"⚠️ Error extrayendo servicios de {pdf_path}: {e}")
        return None


def _find_services_table_start(lineas):
    """Busca el inicio de la tabla de servicios en las líneas del PDF."""
    for i, linea in enumerate(lineas):
        if "Código Unid. Med." in linea and "Descripción" in linea:
            return i
    return None


def _extract_teaching_services(lineas):
    """Extrae servicios de enseñanza que coincidan con el patrón de horas."""
    patron = re.compile(r"^\d{1,2}\s*horas\s+de\s+.*", re.IGNORECASE)
    servicios = []
    
    for linea in lineas:
        linea_clean = linea.strip()
        if patron.match(linea_clean):
            servicios.append(re.sub(r"\s+", " ", linea_clean))
        elif servicios:  # Si ya empezó a encontrar servicios y se corta la secuencia
            break
    
    return servicios


def _format_services_list(servicios):
    """Formatea una lista de servicios con comas y 'y' al final."""
    if len(servicios) > 1:
        return ", ".join(servicios[:-1]) + " y " + servicios[-1]
    else:
        return servicios[0]



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
    """Envía correo con orden de servicio a docente."""
    nombre_formato = nombre.split(",")[0].strip()
    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {CURRENT_YEAR} - {nombre_formato}"
    cuerpo_html = generar_cuerpo_correo_docente_html(mes, CURRENT_YEAR, servicio)
    
    enviar_correo(destinatario, asunto, cuerpo_html, pdf_path, nombre)

def enviar_correo_administrativo(nombre: str, pdf_path: str, destinatario: str, mes: str) -> None:
    """Envía correo con orden de servicio a personal administrativo."""
    asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {CURRENT_YEAR}"
    cuerpo_html = generar_cuerpo_correo_administrativo_html(mes, CURRENT_YEAR)
    
    enviar_correo(destinatario, asunto, cuerpo_html, pdf_path, nombre)

def crear_mapeo_correos(lista_pdfs, nombres_excel, df, tipo_correo="docente"):
    mapeo = {}
    
    for pdf_path in lista_pdfs:
        nombre_extraido = extraer_nombre(pdf_path)
        if not nombre_extraido:
            continue
            
        if nombre_extraido not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre_extraido, nombres_excel)
            if resultado and resultado[1] >= SIMILARITY_THRESHOLD:
                mejor_match = resultado[0]
                fila = df[df['Docente'] == mejor_match]
                
                if not fila.empty:
                    correo = fila['Correo Institucional'].values[0]
                    
                    datos_base = {
                        "pdf_path": pdf_path,
                        "nombre": mejor_match,
                        "correo": correo
                    }
                    
                    if tipo_correo == "docente":
                        # Para docentes, extraer también los servicios
                        servicio = extraer_servicios(pdf_path)
                        datos_base["servicio"] = servicio
                    
                    mapeo[nombre_extraido] = datos_base
    
    return mapeo

def procesar_correos_docente(ruta_excel, hoja, lista_pdfs):
    print("📊 Cargando datos de docentes...")
    
    # Usar repositorio para leer Excel
    if not excel_repo.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el archivo Excel: {ruta_excel}")
    
    df = excel_repo.read_sheet(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    # Validar columnas requeridas
    required_columns = ['Docente', 'Correo Institucional']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Columnas faltantes en Excel: {missing_columns}")

    print(f"✅ Datos cargados: {len(df)} docentes encontrados")
    
    nombres_excel = df['Docente'].astype(str).tolist()

    # Crear mapeo fuzzy optimizado
    print("🔄 Creando mapeo de correos...")
    mapeo_correos = crear_mapeo_correos(lista_pdfs, nombres_excel, df, "docente")

    print("📧 Procesando PDFs para envío...")
    resultados = []
    for pdf_path in lista_pdfs:
        nombre_docente = extraer_nombre(pdf_path)
        if not nombre_docente:
            print(f"⚠️ No se encontró nombre en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_docente in mapeo_correos:
            datos = mapeo_correos[nombre_docente]
            if datos["servicio"]:
                resultados.append(datos)
                print(f"✅ {datos['nombre']} - {datos['correo']} - {datos['servicio']}")
            else:
                print(f"⚠️ No se encontró servicio para {datos['nombre']}.")
        else:
            print(f"⚠️ Coincidencia baja para '{nombre_docente}', omitido.")

    print(f"📋 Procesamiento completado: {len(resultados)} correos listos para enviar")
    return resultados


def procesar_correos_administrativos(ruta_excel, hoja, lista_pdfs):
    print("📊 Cargando datos de personal administrativo...")
    
    # Usar repositorio para leer Excel
    if not excel_repo.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el archivo Excel: {ruta_excel}")
    
    df = excel_repo.read_sheet(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    # Validar columnas requeridas
    required_columns = ['Docente', 'Correo Institucional']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Columnas faltantes en Excel: {missing_columns}")

    print(f"✅ Datos cargados: {len(df)} personas encontradas")
    
    nombres_excel = df['Docente'].astype(str).tolist()

    print("🔄 Creando mapeo de correos...")
    mapeo_correos = crear_mapeo_correos(lista_pdfs, nombres_excel, df, "administrativo")

    print("📧 Procesando PDFs para envío...")
    resultados = []
    for pdf_path in lista_pdfs:
        nombre_administrativo = extraer_nombre(pdf_path)
        if not nombre_administrativo:
            print(f"⚠️ No se encontró nombre en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_administrativo in mapeo_correos:
            datos = mapeo_correos[nombre_administrativo]
            resultados.append(datos)
            print(f"✅ {datos['nombre']} - {datos['correo']}")
        else:
            print(f"⚠️ Coincidencia baja para '{nombre_administrativo}', omitido.")

    print(f"📋 Procesamiento completado: {len(resultados)} correos listos para enviar")
    return resultados