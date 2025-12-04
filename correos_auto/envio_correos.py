import pandas as pd
from PyPDF2 import PdfReader
import re
import datetime
import os
import sys
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from fuzzywuzzy import process

# Gmail API imports
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

año_actual = datetime.datetime.now().year

PALABRAS_PROHIBIDAS = {
    "SERVICIO", "SERVICIOSERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS",
    "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "COORDINACION", "ATENCION",
    "DIGITACION", "CLASIFICACION", "DESCRIPCION", "EVALUACION", "SOLICITUDES",
    "RECEPCION", "ARCHIVO", "VISITAS", "DE", "LA", "Y", "EN", "MES", "DIA", "AÑO",
    "ADJUDICACION", "PROCESO", "SIN", "ADQUISICION", "ORDEN"
}

# Configuración de Gmail API
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

GMAIL_CONFIG = {
    "scopes": [
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.compose',
        'https://www.googleapis.com/auth/gmail.settings.basic'
    ],
    "credentials_file": os.path.join(get_app_dir(), "credentials.json"),
    "token_file": os.path.join(get_app_dir(), "token.pickle"),
    "remitente": "personalcontratado28.flch@unmsm.edu.pe"
}

def autenticar_gmail():
    creds = None
    
    # El archivo token.pickle almacena los tokens de acceso y actualización del usuario.
    if os.path.exists(GMAIL_CONFIG["token_file"]):
        with open(GMAIL_CONFIG["token_file"], 'rb') as token:
            creds = pickle.load(token)
    
    # Si no hay credenciales válidas disponibles, permite al usuario autenticarse.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(GMAIL_CONFIG["credentials_file"]):
                raise FileNotFoundError(
                    f"No se encontró {GMAIL_CONFIG['credentials_file']}. "
                    "Descarga las credenciales OAuth2 desde Google Cloud Console."
                )
            
            flow = InstalledAppFlow.from_client_secrets_file(
                GMAIL_CONFIG["credentials_file"], 
                GMAIL_CONFIG["scopes"]
            )
            creds = flow.run_local_server(port=0)
        
        # Guardar las credenciales para la próxima ejecución
        with open(GMAIL_CONFIG["token_file"], 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)

def obtener_firma_gmail(service):
    try:
        # Intentar obtener la configuración de envío
        send_as = service.users().settings().sendAs().list(userId='me').execute()
        
        for alias in send_as.get('sendAs', []):
            if alias.get('isPrimary'):
                signature = alias.get('signature', '')
                if signature:
                    return signature
        return ""
        
    except Exception as e:
        print(f"⚠️ No se pudo obtener la firma de Gmail: {e}")
        return ""

def _es_nombre_valido(nombre, min_len=3):
    """Valida que un nombre tenga el formato y contenido correcto"""
    if not nombre:
        return False
    partes = nombre.split(',')
    if len(partes) < 2:
        return False
    apellidos = partes[0].strip()
    nombres = partes[1].strip()
    apellidos_letras = len([c for c in apellidos if c.isalpha()])
    nombres_letras = len([c for c in nombres if c.isalpha()])
    return apellidos_letras >= min_len and nombres_letras >= min_len

def _limpiar_nombre(nombre_raw):
    """Limpia y normaliza un nombre extraído"""
    return re.sub(r"\s+", " ", nombre_raw).strip()

def _convertir_a_formato_coma(nombre_sin_coma):
    """Convierte formato 'APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2' a 'APELLIDO1 APELLIDO2, NOMBRE1 NOMBRE2'"""
    primera_palabra = nombre_sin_coma.split()[0] if nombre_sin_coma.split() else ""
    
    if primera_palabra in PALABRAS_PROHIBIDAS:
        return None
        
    palabras = nombre_sin_coma.split()
    if len(palabras) >= 3:
        apellidos = " ".join(palabras[:2])
        nombres = " ".join(palabras[2:])
        return f"{apellidos}, {nombres}"
    return None

def extraer_nombre(pdf_path, debug=False):
    reader = PdfReader(pdf_path)
    texto = ""

    # Concatenar el texto de todas las páginas
    for pagina in reader.pages:
        texto += pagina.extract_text() + "\n"

    candidatos = []

    # Patrón 1: Concepto:UNMSM seguido de números y nombre CON COMA
    match = re.search(r"Concepto:\s*UNMSM\s*\n?\s*\d+\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto, re.IGNORECASE)
    if match:
        nombre_limpio = _limpiar_nombre(match.group(1))
        if _es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 1 (Concepto:UNMSM con coma)", nombre_limpio))

    # Patrón 1b: Concepto:UNMSM seguido de números y nombre SIN COMA
    match = re.search(r"Concepto:\s*UNMSM\s*\n?\s*(\d{8,})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto, re.IGNORECASE)
    if match:
        nombre_formateado = _convertir_a_formato_coma(match.group(2).strip())
        if nombre_formateado and _es_nombre_valido(nombre_formateado, min_len=3):
            candidatos.append(("Patrón 1b (Concepto:UNMSM sin coma)", nombre_formateado))

    # Patrón 2: números seguidos directamente de nombre CON COMA
    match = re.search(r"\d{8,}\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)", texto)
    if match:
        nombre_limpio = _limpiar_nombre(match.group(1))
        if _es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 2 (números con coma)", nombre_limpio))

    # Patrón 2b: números largos seguidos de nombre SIN COMA
    match = re.search(r"\d{8,}\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto)
    if match:
        nombre_formateado = _convertir_a_formato_coma(match.group(1).strip())
        if nombre_formateado and _es_nombre_valido(nombre_formateado, min_len=3):
            candidatos.append(("Patrón 2b (números sin coma)", nombre_formateado))

    # Patrón 2c: Buscar específicamente después de RUC (11 dígitos) seguido del nombre
    match = re.search(r"\bRUC:.*?(\d{11})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})", texto, re.IGNORECASE | re.DOTALL)
    if match:
        nombre_sin_coma = match.group(2).strip()
        if len(nombre_sin_coma.split()) <= 6:  # Límite de palabras
            nombre_formateado = _convertir_a_formato_coma(nombre_sin_coma)
            if nombre_formateado and _es_nombre_valido(nombre_formateado, min_len=3):
                candidatos.append(("Patrón 2c (después de RUC)", nombre_formateado))

    # Patrón 3: nombres con saltos de línea
    match = re.search(r"\b([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\s*,\s*\n?\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\b", texto)
    if match:
        nombre_completo = f"{match.group(1).strip()}, {match.group(2).strip()}"
        nombre_limpio = _limpiar_nombre(nombre_completo)
        if _es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 3 (con saltos)", nombre_limpio))

    # Patrón 4: buscar todos los nombres con formato "APELLIDO(S), NOMBRE(S)"
    matches = re.findall(r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b", texto)
    for match in matches:
        nombre_limpio = _limpiar_nombre(match)
        if _es_nombre_valido(nombre_limpio):
            candidatos.append(("Patrón 4 (fallback)", nombre_limpio))

    if debug and candidatos:
        print("Candidatos encontrados:")
        for patron, nombre in candidatos:
            print(f"  - {patron}: {nombre}")

    # Filtrar y seleccionar el mejor candidato
    if candidatos:
        candidatos_validos = []
        for patron, nombre in candidatos:
            nombre_upper = nombre.upper()
            tiene_prohibida = any(palabra in nombre_upper for palabra in PALABRAS_PROHIBIDAS)
            longitud_ok = 15 <= len(nombre) <= 60
            
            if not tiene_prohibida and longitud_ok:
                candidatos_validos.append((patron, nombre))
        
        if not candidatos_validos:
            candidatos_validos = candidatos
        
        prioridad = {
            "Patrón 1 (Concepto:UNMSM con coma)": 1,
            "Patrón 2 (números con coma)": 2,
            "Patrón 1b (Concepto:UNMSM sin coma)": 3,
            "Patrón 2b (números sin coma)": 4,
            "Patrón 2c (después de RUC)": 5,
            "Patrón 3 (con saltos)": 6,
            "Patrón 4 (fallback)": 7
        }
        
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

    patron = re.compile(r"^\d{1,2}\s*horas\s+de\s+.*", re.IGNORECASE)

    horas = []
    for linea in lineas[idx_inicio:]:
        if patron.match(linea.strip()):
            horas.append(re.sub(r"\s+", " ", linea.strip()))
        elif horas:
            break

    if not horas:
        return None

    if len(horas) > 1:
        return ", ".join(horas[:-1]) + " y " + horas[-1]
    else:
        return horas[0]

def crear_mensaje_gmail_con_firma(service, destinatario: str, asunto: str, cuerpo_html: str, pdf_path: str) -> dict:
    try:
        # Crear mensaje MIME básico (sin firma, Gmail la añadirá)
        msg = MIMEMultipart()
        msg['To'] = destinatario
        msg['Subject'] = asunto
        
        # Añadir cuerpo HTML (SIN firma)
        msg.attach(MIMEText(cuerpo_html, 'html'))
        
        # Adjuntar PDF
        with open(pdf_path, 'rb') as adjunto:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(adjunto.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
            msg.attach(parte)
        
        # Codificar mensaje
        raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        
        # Crear draft primero (esto permite que Gmail añada la firma)
        draft_body = {
            'message': {
                'raw': raw_message
            }
        }
        
        draft = service.users().drafts().create(userId='me', body=draft_body).execute()
        
        # Enviar el draft (Gmail añadirá la firma automáticamente)
        result = service.users().drafts().send(userId='me', body={'id': draft['id']}).execute()
        
        return result
        
    except Exception as e:
        raise Exception(f"Error creando mensaje con firma: {str(e)}")

def generar_cuerpo_correo_docente_html(mes: str, anio: str, servicio: str, firma_html: str = "") -> str:
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

        {firma_html}
    </body>
</html>
"""

def generar_cuerpo_correo_administrativo_html(mes: str, anio: str, firma_html: str = "") -> str:
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

        {firma_html}
    </body>
</html>
"""

def enviar_correo_gmail(service, destinatario: str, asunto: str, cuerpo_html: str, pdf_path: str, nombre: str) -> None:
    """Envía un correo con adjunto PDF usando la API de Gmail"""
    try:
        crear_mensaje_gmail_con_firma(service, destinatario, asunto, cuerpo_html, pdf_path)
        print(f"Correo enviado a {nombre}")
    except HttpError as error:
        print(f"❌ Error enviando correo a {nombre}: {error}")
    except Exception as error:
        print(f"❌ Error inesperado enviando correo a {nombre}: {error}")

def enviar_correo_personalizado(service, nombre: str, pdf_path: str, destinatario: str, mes: str, tipo: str, servicio: str = None) -> None:
    """Función unificada para enviar correos a docentes o administrativos"""
    firma_html = obtener_firma_gmail(service)
    
    if tipo == "docente":
        if not servicio:
            print(f"⚠ No se puede enviar correo a {nombre}: servicio no especificado")
            return
        nombre_formato = nombre.split(",")[0].strip()
        asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {año_actual} - {nombre_formato}"
        cuerpo_html = generar_cuerpo_correo_docente_html(mes, año_actual, servicio, firma_html)
    elif tipo == "administrativo":
        asunto = f"Envío de orden de servicio y solicitud de recibo por honorarios – {mes} {año_actual}"
        cuerpo_html = generar_cuerpo_correo_administrativo_html(mes, año_actual, firma_html)
    else:
        print(f"❌ Tipo de correo no válido: {tipo}")
        return
    
    enviar_correo_gmail(service, destinatario, asunto, cuerpo_html, pdf_path, nombre)

def _crear_mapeo_correos(lista_pdfs, nombres_excel, df, incluir_servicio=True, debug=False):
    """Crea mapeo entre PDFs y datos de correo. Función unificada para docentes y administrativos"""
    mapeo = {}
    
    for pdf_path in lista_pdfs:
        nombre_extraido = extraer_nombre(pdf_path, debug=debug)
        if not nombre_extraido or nombre_extraido in mapeo:
            continue
            
        resultado = process.extractOne(nombre_extraido, nombres_excel)
        if resultado and resultado[1] >= 80:
            mejor_match = resultado[0]
            fila = df[df['Docente'] == mejor_match]
            if not fila.empty:
                correo = fila['Correo Institucional'].values[0]
                datos = {
                    "pdf_path": pdf_path,
                    "nombre": mejor_match,
                    "correo": correo
                }
                
                if incluir_servicio:
                    datos["servicio"] = extraer_servicios(pdf_path)
                    
                mapeo[nombre_extraido] = datos
    
    return mapeo

def _procesar_correos_base(ruta_excel, hoja, lista_pdfs, incluir_servicio=True, debug=False):
    """Función base para procesar correos (unifica lógica para docentes y administrativos)"""
    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    df.columns = df.columns.str.strip()

    if 'Docente' not in df.columns or 'Correo Institucional' not in df.columns:
        raise ValueError("El Excel debe contener columnas 'Docente' y 'Correo Institucional'.")

    nombres_excel = df['Docente'].astype(str).tolist()
    mapeo_correos = _crear_mapeo_correos(lista_pdfs, nombres_excel, df, incluir_servicio, debug)

    resultados = []
    for pdf_path in lista_pdfs:
        nombre_encontrado = extraer_nombre(pdf_path, debug=debug)
        if not nombre_encontrado:
            print(f"⚠ No se encontró nombre válido en {os.path.basename(pdf_path)}, omitido.")
            continue

        if nombre_encontrado in mapeo_correos:
            datos = mapeo_correos[nombre_encontrado]
            if not incluir_servicio or datos.get("servicio"):
                resultados.append(datos)
                servicio_info = f" - {datos['servicio']}" if incluir_servicio else ""
                print(f"{datos['nombre']} - {datos['correo']}{servicio_info}")
            elif incluir_servicio:
                print(f"⚠ No se encontró servicio para {datos['nombre']}.")
        else:
            print(f"⚠ Coincidencia baja para '{nombre_encontrado}' (archivo: {os.path.basename(pdf_path)}), omitido.")

    return resultados

def procesar_correos_docente_gmail(ruta_excel, hoja, lista_pdfs):
    """Procesa correos para docentes"""
    try:
        service = autenticar_gmail()
    except Exception as e:
        print(f"❌ Error autenticando con Gmail API: {e}")
        return []
    
    return _procesar_correos_base(ruta_excel, hoja, lista_pdfs, incluir_servicio=True)

def procesar_correos_administrativos_gmail(ruta_excel, hoja, lista_pdfs, debug=False):
    """Procesa correos para personal administrativo"""
    try:
        service = autenticar_gmail()
    except Exception as e:
        print(f"❌ Error autenticando con Gmail API: {e}")
        return []
    
    return _procesar_correos_base(ruta_excel, hoja, lista_pdfs, incluir_servicio=False, debug=debug)

def enviar_lote_desde_gui(data_para_envio, mes, tipo):
    """Función unificada para enviar lotes de correos desde la GUI"""
    service = autenticar_gmail()
    
    tipo_texto = "docentes" if tipo == "docente" else "personal administrativo"
    print(f"\nEnviando {len(data_para_envio)} correos a {tipo_texto}...")
    
    for datos in data_para_envio:
        servicio = datos.get("servicio") if tipo == "docente" else None
        enviar_correo_personalizado(
            service, 
            datos["nombre"], 
            datos["pdf_path"], 
            datos["correo"], 
            mes, 
            tipo,
            servicio
        )

# Funciones de compatibilidad para mantener la API existente
def enviar_lote_desde_gui_docentes(data_para_envio, mes):
    """Wrapper para mantener compatibilidad con código existente"""
    enviar_lote_desde_gui(data_para_envio, mes, "docente")

def enviar_lote_desde_gui_administrativos(data_para_envio, mes):
    """Wrapper para mantener compatibilidad con código existente"""
    enviar_lote_desde_gui(data_para_envio, mes, "administrativo")

# Funciones de conveniencia para migración desde SMTP
def enviar_lote_docentes_gmail(ruta_excel, hoja, lista_pdfs, mes):
    """Función completa para procesar y enviar correos a docentes"""
    resultados = procesar_correos_docente_gmail(ruta_excel, hoja, lista_pdfs)
    enviar_lote_desde_gui_docentes(resultados, mes)

def enviar_lote_administrativos_gmail(ruta_excel, hoja, lista_pdfs, mes):
    """Función completa para procesar y enviar correos a personal administrativo"""
    resultados = procesar_correos_administrativos_gmail(ruta_excel, hoja, lista_pdfs)
    enviar_lote_desde_gui_administrativos(resultados, mes)