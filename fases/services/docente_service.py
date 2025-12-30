import os
from typing import Tuple
from fases.models import DocenteData, DocumentConfig


class DocenteService:
    
    @staticmethod
    def crear_carpeta_docente(carpeta_base: str, config: DocumentConfig, docente: DocenteData) -> str:
        carpeta_fase = os.path.join(carpeta_base, config.obtener_nombre_carpeta_fase())
        carpeta_docente = os.path.join(carpeta_fase, docente.nombre_limpio)
        os.makedirs(carpeta_docente, exist_ok=True)
        return carpeta_docente
    
    @staticmethod
    def obtener_datos_contacto_formateados(docente: DocenteData) -> dict:
        return {
            "direccion_cot": f"Dirección: {docente.direccion}",
            "ruc_docente_cot": f"RUC N.º {docente.ruc}",
            "correo_docente_cot": f"Correo: {docente.correo}",
            "celular_cot": f"Teléfono: {docente.celular}",
            "dni_cot": f"DNI: {docente.dni_formateado}"
        }
    
    @staticmethod
    def validar_docente(docente: DocenteData) -> Tuple[bool, str]:
        if not docente.nombre or docente.nombre == "N/A":
            return False, "Nombre del docente no válido"
        
        # DNI es opcional, solo advertir si no está presente
        if not docente.dni:
            print(f"⚠️  Advertencia: {docente.nombre} - DNI no proporcionado")
        
        if docente.estado_docente not in ["CONTRATO", "TERCERO"]:
            return False, f"{docente.nombre} - Estado de docente inválido: {docente.estado_docente}"
        
        return True, ""
    
    @staticmethod
    def obtener_ruta_firma(docente: DocenteData, config: DocumentConfig) -> str:
        return config.obtener_ruta_firma(docente.nombre_limpio)
