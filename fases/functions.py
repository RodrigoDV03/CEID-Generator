import pandas as pd
import os
import sys
import re
from docx import Document
from decimal import Decimal, ROUND_HALF_UP
from num2words import num2words

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", nombre)


def limpiar_numero(valor):
    return "" if pd.isna(valor) else str(valor).split('.')[0]


def reemplazar_en_parrafos(documento, reemplazos):
    for parrafo in documento.paragraphs:
        for marcador, valor in reemplazos.items():
            if marcador in parrafo.text:
                texto_nuevo = parrafo.text.replace(marcador, valor)
                # Limpiar runs existentes
                for run in parrafo.runs:
                    run.text = ''
                # Establecer nuevo texto
                if parrafo.runs:
                    parrafo.runs[0].text = texto_nuevo


def reemplazar_en_tablas(documento, reemplazos):
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for marcador, valor in reemplazos.items():
                        if marcador in parrafo.text:
                            texto_nuevo = parrafo.text.replace(marcador, valor)
                            # Limpiar runs existentes
                            for run in parrafo.runs:
                                run.text = ''
                            # Establecer nuevo texto
                            if parrafo.runs:
                                parrafo.runs[0].text = texto_nuevo

def monto_a_letras(monto):
    try:
        monto = Decimal(str(monto)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        entero = int(monto)
        centavos = int((monto - Decimal(entero)) * 100)
        return f"{num2words(entero, lang='es')} y {centavos:02d}/100 soles"
    except Exception:
        return "N/A"


def redactar_cursos(cadena):
    if not isinstance(cadena, str):
        return "N/A"
    
    cursos = [c.strip() for c in cadena.split("/") if c.strip()]
    if not cursos:
        return "N/A"
    
    # Construir descripción de servicios
    resultado = f"servicio de dictado de 28 horas de clases de {cursos[0]}"
    for curso in cursos[1:]:
        resultado += f", 28 horas de clases de {curso}"
    
    return resultado

def ruta_absoluta_relativa(path_relativo):
    if getattr(sys, 'frozen', False):
        # Ejecutable empaquetado con PyInstaller
        base_path = sys._MEIPASS
    else:
        # Desarrollo normal
        base_path = os.path.abspath(".")
    return os.path.join(base_path, path_relativo)


def generar_documento(modelo_path, reemplazos, ruta_salida, firma_path=None):
    try:
        if not modelo_path or not os.path.exists(modelo_path):
            print(f"⚠️ Plantilla no encontrada: {modelo_path}")
            return False

        # Cargar documento plantilla
        doc = Document(modelo_path)

        # Reemplazar texto (excluyendo marcador de firma)
        reemplazos_texto = {k: v for k, v in reemplazos.items() if k != "firma_docente"}
        reemplazar_en_parrafos(doc, reemplazos_texto)
        reemplazar_en_tablas(doc, reemplazos_texto)
        
        # Insertar firma si está disponible
        if firma_path and os.path.exists(firma_path):
            _insertar_firma_en_documento(doc, firma_path)

        # Crear directorio si no existe
        os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
        
        # Guardar documento
        doc.save(ruta_salida)
        return True
        
    except Exception as e:
        print(f"⚠️ Error generando documento: {e}")
        return False


def _insertar_firma_en_documento(documento, firma_path):
    # Insertar en párrafos
    for parrafo in documento.paragraphs:
        if "firma_docente" in parrafo.text:
            parrafo.text = ""
            run = parrafo.add_run()
            run.add_picture(firma_path)

    # Insertar en tablas
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if "firma_docente" in parrafo.text:
                        parrafo.text = ""
                        run = parrafo.add_run()
                        run.add_picture(firma_path)