from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Cache para estilos de Excel para evitar recrearlos
_excel_styles_cache = {}

def get_excel_style(style_name):
    if style_name not in _excel_styles_cache:
        
        if style_name == "thin_border":
            _excel_styles_cache[style_name] = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        elif style_name == "header_fill":
            _excel_styles_cache[style_name] = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        elif style_name == "header_font":
            _excel_styles_cache[style_name] = Font(bold=True, color="ffffff", size=12)
        elif style_name == "title_font":
            _excel_styles_cache[style_name] = Font(bold=True, color="ffffff", size=22)
        elif style_name == "bold_font":
            _excel_styles_cache[style_name] = Font(bold=True)
        elif style_name == "center_alignment":
            _excel_styles_cache[style_name] = Alignment(horizontal='center', vertical='center')
        elif style_name == "title_alignment":
            _excel_styles_cache[style_name] = Alignment(horizontal='center', vertical='center', wrap_text=True)
        elif style_name == "money_format":
            _excel_styles_cache[style_name] = '"S/ "#,##0.00'
            
    return _excel_styles_cache[style_name]

def aplicar_formato_excel_optimizado(ws, max_col, titulo_fusionado, es_planilla=False):
    ws.insert_rows(1)
    rango_titulo = f"A1:{get_column_letter(max_col)}1"
    ws.merge_cells(rango_titulo)
    ws["A1"].value = titulo_fusionado
    ws["A1"].alignment = get_excel_style("title_alignment")
    ws["A1"].fill = get_excel_style("header_fill")
    ws["A1"].font = get_excel_style("title_font")

    thin_border = get_excel_style("thin_border")
    header_fill = get_excel_style("header_fill")
    header_font = get_excel_style("header_font")
    center_alignment = get_excel_style("center_alignment")
    bold_font = get_excel_style("bold_font")
    money_format = get_excel_style("money_format")

    header_cells = [ws.cell(row=2, column=col) for col in range(1, max_col + 1)]
    for celda in header_cells:
        celda.alignment = center_alignment
        celda.fill = header_fill
        celda.font = header_font

    max_row = ws.max_row
    for row_cells in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row_cells:
            cell.border = thin_border
            cell.alignment = center_alignment

    if es_planilla:
        fila_total = max_row + 1
        ws.merge_cells(f"A{fila_total}:G{fila_total}")
        celda_total = ws[f"A{fila_total}"]
        celda_total.alignment = center_alignment
        celda_total.fill = header_fill

        # Columnas a sumar - aplicar en lote
        columnas_sumar = ['H', 'I', 'J', 'K', 'L', 'M']
        
        for col in columnas_sumar:
            celda = ws[f"{col}{fila_total}"]
            celda.value = f"=SUM({col}3:{col}{fila_total-1})"
            celda.font = bold_font
            celda.alignment = center_alignment
            celda.border = thin_border

        columnas_moneda = ['E', 'H', 'K', 'L', 'M']
        
        for col in columnas_moneda:
            # Aplicar formato a todo el rango de una vez
            for row_num in range(3, fila_total + 1):
                ws[f"{col}{row_num}"].number_format = money_format

    ajustar_anchos_columnas_optimizado(ws, max_col)

def ajustar_anchos_columnas_optimizado(ws, max_col):
    anchos_predefinidos = {
        1: 5,   # N°
        2: 30,  # Docente
        3: 15,  # Sede
        4: 8,   # Categoria (Letra)
        5: 12,  # Categoria (Monto)
        6: 15,  # N°. Ruc
        7: 40,  # Curso
        8: 12,  # Curso Dictado
        9: 10,  # Extra Curso
        10: 12, # Cantidad Cursos
        11: 15, # Diseño de Examenes
        12: 12, # Examen Clasif.
        13: 15, # Total Pago S/.
        14: 10  # Estado
    }
    
    for col in range(1, min(max_col + 1, len(anchos_predefinidos) + 1)):
        column_letter = get_column_letter(col)
        if col in anchos_predefinidos:
            ws.column_dimensions[column_letter].width = anchos_predefinidos[col]
        else:
            # Ancho por defecto para columnas adicionales
            ws.column_dimensions[column_letter].width = 12

def procesar_formato_multiple_hojas(wb, hojas_con_titulo, numero_carga_letra, month):
    for hoja, titulo_fusionado in hojas_con_titulo:
        if hoja in wb.sheetnames:
            ws = wb[hoja]
            max_col = ws.max_column
            
            es_planilla = (hoja == f"{numero_carga_letra} Planilla {month}" or hoja == "Planilla consolidada")
            
            aplicar_formato_excel_optimizado(ws, max_col, titulo_fusionado, es_planilla)