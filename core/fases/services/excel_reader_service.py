import pandas as pd
from typing import List, Any
from core.fases.models import DocenteData, PaymentData, CursoDetalle


class ExcelReaderService:

    @staticmethod
    def _normalizar_texto(valor: Any) -> str:
        if pd.isna(valor):
            return ""
        return " ".join(str(valor).strip().upper().split())

    @staticmethod
    def _buscar_columna(fila: Any, aliases: List[str], default: Any = "") -> Any:
        for col_name in aliases:
            try:
                valor = getattr(fila, col_name)
                if pd.notna(valor):
                    valor_str = str(valor).strip()
                    if valor_str and valor_str.lower() != 'nan':
                        return valor
            except AttributeError:
                continue
        return default

    @staticmethod
    def _parse_float_safe(valor: Any) -> float:
        if pd.isna(valor):
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)

        valor_str = str(valor).strip()
        if not valor_str:
            return 0.0

        # Soportar formatos tipo "1,234.56" y "1234,56"
        valor_str = valor_str.replace(" ", "")
        if "," in valor_str and "." in valor_str:
            valor_str = valor_str.replace(",", "")
        elif "," in valor_str and "." not in valor_str:
            valor_str = valor_str.replace(",", ".")

        try:
            return float(valor_str)
        except ValueError:
            return 0.0
    
    @staticmethod
    def leer_planilla(ruta_excel: str, nombre_hoja: str) -> pd.DataFrame:
        try:
            # Especificar dtype para preservar ceros iniciales en número de contrato
            columnas_contrato = [
                'Nro_Contrato',
                'Nro_contrato',
                'N° Contrato',
                'Numero de contrato',
                'Número de contrato'
            ]
            dtype_specs = {col: str for col in columnas_contrato}
            converters = {
                col: (lambda x: str(x).strip() if pd.notna(x) else '')
                for col in columnas_contrato
            }
            df = pd.read_excel(
                ruta_excel, 
                sheet_name=nombre_hoja,
                dtype=dtype_specs,
                converters=converters
            )
            df.columns = df.columns.str.strip()
            return df
        except FileNotFoundError:
            raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel}")
        except Exception as e:
            raise ValueError(f"Error al leer la hoja '{nombre_hoja}': {str(e)}")
    
    @staticmethod
    def leer_control_pagos(ruta_excel: str) -> pd.DataFrame:
        # Detectar automáticamente la fila de encabezados para tolerar plantillas
        # con títulos en la primera o segunda fila.
        df_preview = pd.read_excel(ruta_excel, sheet_name=0, header=None)

        posibles_encabezados = {
            'APELLIDOS Y NOMBRES',
            'DOCENTE',
            'MONTO TOTAL PARA CONTRATO S/',
            'TOTAL',
            'PRIMERA ARMADA'
        }

        header_idx = 0
        max_scan = min(len(df_preview), 10)
        for i in range(max_scan):
            row_vals = {
                ExcelReaderService._normalizar_texto(v)
                for v in df_preview.iloc[i].tolist()
            }
            if row_vals & posibles_encabezados:
                header_idx = i
                break

        # Forzar que columnas de números de contrato se lean como texto
        dtype_dict = {
            'Numero de contrato': str,
            'Número de contrato': str,
            'N° Contrato': str,
            'Nro_Contrato': str,
            'Nro_contrato': str
        }
        df = pd.read_excel(ruta_excel, sheet_name=0, header=header_idx, dtype=dtype_dict)
        df.columns = df.columns.str.strip()
        return df
    
    @staticmethod
    def obtener_hojas_disponibles(ruta_excel: str) -> List[str]:
        return [str(sheet) for sheet in pd.ExcelFile(ruta_excel).sheet_names]
    
    @staticmethod
    def _limpiar_numero(valor: Any) -> str:
        if pd.isna(valor):
            return ""
        return str(valor).split('.')[0]

    @staticmethod
    def _construir_curso_resumido(fila: Any) -> str:
        cursos = []
        for col_name in ["Curso", "Curso_Virtual", "Curso_Presencial"]:
            valor = getattr(fila, col_name, "")
            if pd.isna(valor):
                continue
            valor_str = str(valor).strip()
            if valor_str and valor_str.lower() != 'nan':
                cursos.append(valor_str)
        return " / ".join(cursos)
    
    @staticmethod
    def extraer_docente_data(fila: Any) -> DocenteData:
        # Extraer número de contrato con múltiples variaciones de nombre
        numero_contrato = ""
        for col_name in ["Nro_Contrato", "Nro_contrato", "N° Contrato", "Numero de contrato", "Número de contrato"]:
            try:
                numero_contrato = getattr(fila, col_name, "")
                if numero_contrato and numero_contrato != "":
                    break
            except AttributeError:
                continue
        
        curso = str(getattr(fila, "Curso", "")).strip()
        if not curso:
            curso = ExcelReaderService._construir_curso_resumido(fila)

        return DocenteData(
            nombre=str(getattr(fila, "Docente", "N/A")),
            dni=ExcelReaderService._limpiar_numero(getattr(fila, "Numero_dni", "")),
            ruc=ExcelReaderService._limpiar_numero(getattr(fila, "N_Ruc", "")),
            direccion=str(getattr(fila, "Domicilio_docente", "")).strip(),
            correo=str(getattr(fila, "Correo_personal", "")),
            celular=ExcelReaderService._limpiar_numero(getattr(fila, "Numero_celular", "")),
            curso=curso,
            categoria_letra=str(getattr(fila, "Categoria_letra", "")).strip().upper(),
            formacion_academica=str(getattr(fila, "Formacion_academica", "")),
            experiencia_laboral=str(getattr(fila, "Experiencia_laboral", "")),
            requisitos_adicional=str(getattr(fila, "Requisitos_adicional", "")),
            finalidad_publica=str(getattr(fila, "Finalidad_publica", "")),
            especialidad=str(getattr(fila, "Especialidad", "")),
            actividades_admin=str(getattr(fila, "Actividades_admin", "")),
            estado_docente=str(getattr(fila, "Estado_docente", "TERCERO")).strip().upper(),
            numero_contrato=ExcelReaderService._extraer_numero_contrato(numero_contrato),
            idioma = str(getattr(fila, "Docente_idioma", "")),
            modalidad = str(getattr(fila, "Modalidad", ""))
        )
    
    @staticmethod
    def extraer_payment_data(fila: Any) -> PaymentData:
        categoria_valor = getattr(fila, "Categoria_monto", 1)
        if pd.isna(categoria_valor):
            categoria_valor = 1.0
        
        # Leer valores con manejo robusto
        total_pago = float(getattr(fila, "Total_pago", getattr(fila, "Monto_total_pagar", 0)))
        servicio_act = float(getattr(fila, "Servicio_actualizacion", 0))
        bono = float(getattr(fila, "Bono", 0))
        disenio = float(getattr(fila, "Disenio_examenes", 0))
        examen_clasif = float(getattr(fila, "Examen_clasif", 0))
        
        # Debug para casos problemáticos
        if total_pago == 0 and (servicio_act > 0 or disenio > 0 or examen_clasif > 0):
            print(f"⚠️ Total_pago es 0 pero hay otros montos: Servicio={servicio_act}, Diseño={disenio}, Examen={examen_clasif}")
        
        return PaymentData(
            categoria_monto=float(categoria_valor),
            total_pago=total_pago,
            servicio_actualizacion=servicio_act,
            bono=bono,
            disenio_examenes=disenio,
            examen_clasificacion=examen_clasif
        )
    
    @staticmethod
    def extraer_payment_data_control(fila: Any) -> PaymentData:
        monto_total = ExcelReaderService._buscar_columna(
            fila,
            ["MONTO TOTAL PARA CONTRATO S/", "TOTAL", "Monto total", "MONTO TOTAL"],
            0
        )
        primera = ExcelReaderService._buscar_columna(fila, ["Primera armada", "PRIMERA ARMADA"], 0)
        segunda = ExcelReaderService._buscar_columna(fila, ["Segunda armada", "SEGUNDA ARMADA"], 0)
        tercera = ExcelReaderService._buscar_columna(fila, ["Tercera armada", "TERCERA ARMADA"], 0)

        return PaymentData(
            monto_total_contrato=ExcelReaderService._parse_float_safe(monto_total),
            primera_armada=ExcelReaderService._parse_float_safe(primera),
            segunda_armada=ExcelReaderService._parse_float_safe(segunda),
            tercera_armada=ExcelReaderService._parse_float_safe(tercera)
        )
    
    @staticmethod
    def extraer_docente_nombre_control(fila: Any) -> str:
        nombre = ExcelReaderService._buscar_columna(
            fila,
            ["APELLIDOS Y NOMBRES", "DOCENTE", "Docente", "NOMBRES Y APELLIDOS"],
            "N/A"
        )
        return str(nombre).strip()
    
    @staticmethod
    def extraer_numero_contrato_control(fila: Any) -> str:
        # Intentar con diferentes variaciones de nombre de columna
        numero = ""
        for col_name in ["Numero de contrato", "Número de contrato", "N° Contrato", "Nro_Contrato", "Nro_contrato"]:
            try:
                numero = getattr(fila, col_name, "")
                if numero and numero != "":
                    break
            except AttributeError:
                continue
        return ExcelReaderService._extraer_numero_contrato(numero)
    
    @staticmethod
    def _extraer_numero_contrato(valor: Any) -> str:
        """Extrae y formatea el número de contrato preservando ceros iniciales."""
        if pd.isna(valor) or valor == "" or valor == "N/A":
            return "001"  # Valor por defecto si no hay número
        
        valor_str = str(valor).strip()
        
        # Si contiene punto decimal (ej: "50.0" de Excel), removerlo
        if '.' in valor_str:
            try:
                # Convertir a float y luego a int para remover decimales
                valor_num = int(float(valor_str))
                return str(valor_num).zfill(max(4, len(str(valor_num))))
            except (ValueError, TypeError):
                pass

        if valor_str.isdigit():
            return valor_str.zfill(max(4, len(valor_str)))
        
        # Retornar como string limpio si no hay decimales
        return valor_str if valor_str else "001"
    
    @staticmethod
    def leer_cursos_detallados_por_docente(
        ruta_excel: str,
        nombre_hoja: str,
        nombre_docente: str
    ) -> List[CursoDetalle]:
        """
        Lee la hoja Planilla_Generador y extrae todos los cursos/servicios
        de un docente específico, tanto desde formato expandido como resumido.
        
        Args:
            ruta_excel: Ruta al archivo Excel de la planilla
            nombre_hoja: Nombre de la hoja (generalmente "Planilla_Generador")
            nombre_docente: Nombre del docente a buscar
        
        Returns:
            Lista de objetos CursoDetalle con todos los cursos/servicios del docente
        """
        try:
            df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja)
            df.columns = df.columns.str.strip()
            
            # Validar que existan las columnas necesarias para cursos detallados
            columnas_requeridas = ['Curso_Individual', 'Modalidad_Curso', 'Tipo_Servicio', 'Horas_Servicio', 'Monto_Individual']
            columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
            
            if columnas_faltantes:
                filas_docente = df[df['Docente'].str.strip().str.upper() == nombre_docente.strip().upper()]
                if filas_docente.empty:
                    return []

                fila = filas_docente.iloc[0]
                cursos = []
                categoria_monto = float(fila.get('Categoria_monto', 0) or 0)

                for nombre_columna, modalidad in [('Curso_Virtual', 'VIRTUAL'), ('Curso_Presencial', 'INPERSON')]:
                    cursos_raw = fila.get(nombre_columna, '')
                    if pd.isna(cursos_raw):
                        continue
                    for curso_nombre in [c.strip() for c in str(cursos_raw).split('/') if c.strip()]:
                        cursos.append(CursoDetalle(
                            nombre=curso_nombre,
                            modalidad=modalidad,
                            tipo_servicio='CURSO_DICTADO',
                            horas=28,
                            monto=float(categoria_monto * 28)
                        ))

                examen_clasif = float(fila.get('Examen_clasif', 0) or 0)
                if examen_clasif > 0:
                    horas = int(round(examen_clasif / categoria_monto)) if categoria_monto > 0 else 0
                    cursos.append(CursoDetalle(
                        nombre='Examen de clasificación',
                        modalidad='VIRTUAL',
                        tipo_servicio='EXAMEN_CLASIF',
                        horas=horas,
                        monto=examen_clasif
                    ))

                disenio_examenes = float(fila.get('Disenio_examenes', 0) or 0)
                if disenio_examenes > 0:
                    horas = int(round(disenio_examenes / categoria_monto)) if categoria_monto > 0 else 0
                    cursos.append(CursoDetalle(
                        nombre='Diseño de exámenes',
                        modalidad='N/A',
                        tipo_servicio='DISENO_EXAMENES',
                        horas=horas,
                        monto=disenio_examenes
                    ))

                servicio_actualizacion = float(fila.get('Servicio_actualizacion', 0) or 0)
                if servicio_actualizacion > 0:
                    horas_total = int(float(fila.get('Horas_Total', 0) or 0))
                    cursos.append(CursoDetalle(
                        nombre='Servicio de actualización de materiales de enseñanza',
                        modalidad='VIRTUAL',
                        tipo_servicio='SERVICIO_ACTUALIZACION',
                        horas=horas_total,
                        monto=servicio_actualizacion
                    ))

                bono = float(fila.get('Bono', 0) or 0)
                if bono > 0:
                    cursos.append(CursoDetalle(
                        nombre='Servicio de diseño y evaluación del examen anual',
                        modalidad='N/A',
                        tipo_servicio='BONO',
                        horas=0,
                        monto=bono
                    ))

                return cursos
            
            # Filtrar solo las filas de este docente
            filas_docente = df[df['Docente'].str.strip().str.upper() == nombre_docente.strip().upper()]
            
            cursos = []
            for _, fila in filas_docente.iterrows():
                try:
                    curso = CursoDetalle(
                        nombre=str(fila['Curso_Individual']),
                        modalidad=str(fila['Modalidad_Curso']),
                        tipo_servicio=str(fila['Tipo_Servicio']),
                        horas=int(fila['Horas_Servicio']) if not pd.isna(fila['Horas_Servicio']) else 0,
                        monto=float(fila['Monto_Individual']) if not pd.isna(fila['Monto_Individual']) else 0.0
                    )
                    cursos.append(curso)
                except Exception as e:
                    print(f"⚠️ Error al procesar curso: {e}")
                    continue
            
            return cursos
            
        except Exception as e:
            # Error silencioso, retornar lista vacía para usar fallback
            return []
