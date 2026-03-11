import pandas as pd
from typing import List, Any
from fases.models import DocenteData, PaymentData, CursoDetalle


class ExcelReaderService:
    
    @staticmethod
    def leer_planilla(ruta_excel: str, nombre_hoja: str) -> pd.DataFrame:
        try:
            # Forzar que columnas de números de contrato se lean como texto (string)
            # para preservar ceros adelante (ej: 0057, 0130)
            dtype_dict = {
                'Nro_Contrato': str,
                'Nro_contrato': str,
                'N° Contrato': str,
                'Numero de contrato': str,
                'Número de contrato': str
            }
            df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja, dtype=dtype_dict)
            df.columns = df.columns.str.strip()
            return df
        except FileNotFoundError:
            raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel}")
        except Exception as e:
            raise ValueError(f"Error al leer la hoja '{nombre_hoja}': {str(e)}")
    
    @staticmethod
    def leer_control_pagos(ruta_excel: str) -> pd.DataFrame:
        # Forzar que columnas de números de contrato se lean como texto
        dtype_dict = {
            'Numero de contrato': str,
            'Número de contrato': str,
            'N° Contrato': str,
            'Nro_Contrato': str,
            'Nro_contrato': str
        }
        df = pd.read_excel(ruta_excel, sheet_name=0, header=1, dtype=dtype_dict)
        df.columns = df.columns.str.strip()
        return df
    
    @staticmethod
    def obtener_hojas_disponibles(ruta_excel: str) -> List[str]:
        return pd.ExcelFile(ruta_excel).sheet_names
    
    @staticmethod
    def _limpiar_numero(valor: Any) -> str:
        if pd.isna(valor):
            return ""
        return str(valor).split('.')[0]
    
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
        
        return DocenteData(
            nombre=str(getattr(fila, "Docente", "N/A")),
            dni=ExcelReaderService._limpiar_numero(getattr(fila, "Numero_dni", "")),
            ruc=ExcelReaderService._limpiar_numero(getattr(fila, "N_Ruc", "")),
            direccion=str(getattr(fila, "Domicilio_docente", "")).strip(),
            correo=str(getattr(fila, "Correo_personal", "")),
            celular=ExcelReaderService._limpiar_numero(getattr(fila, "Numero_celular", "")),
            curso=str(getattr(fila, "Curso", "")),
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
        
        return PaymentData(
            categoria_monto=float(categoria_valor),
            total_pago=float(getattr(fila, "Total_pago", 0)),
            servicio_actualizacion=float(getattr(fila, "Servicio_actualizacion", 0)),
            bono=float(getattr(fila, "Bono", 0)),
            disenio_examenes=float(getattr(fila, "Disenio_examenes", 0)),
            examen_clasificacion=float(getattr(fila, "Examen_clasif", 0))
        )
    
    @staticmethod
    def extraer_payment_data_control(fila: Any) -> PaymentData:
        return PaymentData(
            monto_total_contrato=float(getattr(fila, "MONTO TOTAL PARA CONTRATO S/", 0)),
            primera_armada=float(getattr(fila, "Primera armada", 0)),
            segunda_armada=float(getattr(fila, "Segunda armada", 0)),
            tercera_armada=float(getattr(fila, "Tercera armada", 0))
        )
    
    @staticmethod
    def extraer_docente_nombre_control(fila: Any) -> str:
        return str(getattr(fila, "APELLIDOS Y NOMBRES", "N/A")).strip()
    
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
        """Extrae y formatea el número de contrato preservando ceros adelante."""
        if pd.isna(valor) or valor == "" or valor == "N/A":
            return ""  # Retornar string vacío si no hay valor
        
        # Convertir a string y limpiar espacios
        valor_str = str(valor).strip()
        
        # Eliminar ".0" si pandas lo agregó (ej: "0057.0" -> "0057")
        if valor_str.endswith('.0'):
            valor_str = valor_str[:-2]
        
        return valor_str if valor_str else ""
    
    @staticmethod
    def leer_cursos_detallados_por_docente(
        ruta_excel: str,
        nombre_hoja: str,
        nombre_docente: str
    ) -> List[CursoDetalle]:
        """
        Lee la hoja Planilla_Generador expandida y extrae todos los cursos/servicios
        de un docente específico con su modalidad.
        
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
                # Silenciosamente retornar lista vacía si no tiene la estructura expandida
                return []
            
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
