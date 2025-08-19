import pandas as pd
import os
from fuzzywuzzy import process
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def _norm(s):
    if not isinstance(s, str):
        return ""
    return " ".join(s.strip().upper().split())

def actualizar_control_pagos(planilla_path, control_path, numero_armada):
    for path in [planilla_path, control_path]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"No se encontró el archivo: {path}")
        if not path.lower().endswith((".xls", ".xlsx")):
            raise ValueError(f"Extensión inválida (solo .xls o .xlsx): {path}")

    # Lectura de planilla con pandas (solo para extraer montos)
    df_planilla = pd.read_excel(planilla_path, sheet_name="Planilla_Generador")

    # Mapeo de armadas
    columnas_armada = {
        "Primera": "Primera armada",
        "Segunda": "Segunda armada",
        "Tercera": "Tercera armada",
    }

    if numero_armada not in columnas_armada:
        raise ValueError("Número de armada inválido. Use 'Primera', 'Segunda' o 'Tercera'.")
    col_armada_name = columnas_armada[numero_armada]

    # Matching Docente y Subtotal_Pago
    nombres_planilla = df_planilla["Docente"].dropna().tolist()

    def monto_por_nombre(docente_control):
        if not isinstance(docente_control, str):
            return None
        match = process.extractOne(docente_control, nombres_planilla)
        if not match:
            return None
        nombre_match, score = match
        if score < 85:
            return None
        fila = df_planilla.loc[df_planilla["Docente"] == nombre_match, "Subtotal_pago"]
        return None if fila.empty else float(fila.values[0])

    wb = load_workbook(control_path)
    ws = wb.active
    # "APELLIDOS Y NOMBRES"
    header_row = None
    for r in range(1, min(ws.max_row, 15) + 1):
        values = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        if any(v.upper() == "APELLIDOS Y NOMBRES" for v in values):
            header_row = r
            break
    if header_row is None:
        raise RuntimeError("No se encontró la fila de encabezados con 'APELLIDOS Y NOMBRES'.")

    # Armamos un índice {titulo: índice_columna(1-based)}
    headers = {}
    for cell in ws[header_row]:
        key = _norm(cell.value)
        if key:
            headers[key] = cell.column  # 1-based

    # Nombres de cabecera tal cual están en el archivo
    req = [
        "APELLIDOS Y NOMBRES",
        "MONTO TOTAL PARA CONTRATO S/",
        "PRIMERA ARMADA",
        "SEGUNDA ARMADA",
        "TERCERA ARMADA",
        col_armada_name.upper(),
        "SALDO RESTANTE",
    ]
    for titulo in req:
        if _norm(titulo) not in headers:
            raise RuntimeError(f"No se encontró la columna: {titulo}")

    col_idx_nombre  = headers["APELLIDOS Y NOMBRES"]
    col_idx_total   = headers["MONTO TOTAL PARA CONTRATO S/"]
    col_idx_primera = headers["PRIMERA ARMADA"]
    col_idx_segunda = headers["SEGUNDA ARMADA"]
    col_idx_tercera = headers["TERCERA ARMADA"]
    col_idx_armada  = headers[col_armada_name.upper()]
    col_idx_saldo   = headers["SALDO RESTANTE"]

    # Letras de columna para armar fórmula
    L_total   = get_column_letter(col_idx_total)
    L_primera = get_column_letter(col_idx_primera)
    L_segunda = get_column_letter(col_idx_segunda)
    L_tercera = get_column_letter(col_idx_tercera)

    first_data_row = header_row + 1
    for r in range(first_data_row, ws.max_row + 1):
        nombre = ws.cell(row=r, column=col_idx_nombre).value
        if not nombre or str(nombre).strip().upper() in {"APELLIDOS Y NOMBRES", ""}:
            continue

        monto = monto_por_nombre(str(nombre))
        if monto is None:
            continue  # sin coincidencia

        ws.cell(row=r, column=col_idx_armada).value = float(monto)

        formula = f"={L_total}{r}-SUM({L_primera}{r},{L_segunda}{r},{L_tercera}{r})"
        ws.cell(row=r, column=col_idx_saldo).value = formula

    wb.save(control_path)
    print(f"Archivo actualizado guardado en: {control_path}")
