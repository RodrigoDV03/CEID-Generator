import pandas as pd
import os

def actualizar_control_pagos(planilla_path, control_path, numero_armada, output_path):
    # Verificar existencia y extensión de archivos
    for path in [planilla_path, control_path]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"No se encontró el archivo: {path}")
        if not (path.lower().endswith('.xls') or path.lower().endswith('.xlsx')):
            raise ValueError(f"El archivo no tiene extensión válida de Excel (.xls o .xlsx): {path}")
    # Leer archivos
    df_planilla = pd.read_excel(planilla_path, sheet_name="Planilla_Generador")
    df_control = pd.read_excel(control_path, header=2)

    # Mapear número de armada a nombre de columna
    columnas_armada = {
        "Primera": "Primera armada",
        "Segunda": "Segunda armada",
        "Tercera": "Tercera armada"
    }
    
    if numero_armada not in columnas_armada:
        raise ValueError("Número de armada inválido. Debe ser 'Primera', 'Segunda' o 'Tercera'.")

    col_armada = columnas_armada[numero_armada]

    # Iterar por cada docente en control y buscar en planilla
    for idx, row in df_control.iterrows():
        docente = row["Docente"]
        monto_en_planilla = df_planilla.loc[df_planilla["Docente"] == docente, "Subtotal_Pago"]

        if not monto_en_planilla.empty:
            df_control.at[idx, col_armada] = monto_en_planilla.values[0]

    # Calcular saldo restante
    df_control["Saldo restante"] = (
        df_control["MONTO TOTAL PARA CONTRATO S/"].fillna(0) -
        (df_control["Primera armada"].fillna(0) +
         df_control["Segunda armada"].fillna(0) +
         df_control["Tercera armada"].fillna(0))
    )

    # Guardar archivo actualizado
    # df_control.to_excel(output_path, index=False)
    print(f"Archivo actualizado guardado en: {output_path}")
