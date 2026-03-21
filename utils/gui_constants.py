from datetime import datetime

# COLORES BASE
PRIMARY_COLOR = "#1d5fb6"        # azul principal
PRIMARY_DARK = "#114b98"         # fondo
ACCENT_COLOR = "#f38b1a"         # naranja
ACCENT_HOVER = "#e8aa69"

BG_COLOR = "#acd0ff"             # fondo general claro
CARD_COLOR = "#ffffff"
CONSOLE_BG = "#2c2f33"

TEXT_COLOR = "#1f2937"
TEXT_LIGHT = "#6b7280"
WHITE_COLOR = "#ffffff"

# CONSERVAMOS COMPATIBILIDAD
SECTION_COLOR = PRIMARY_COLOR
BUTTON_BG_COLOR = ACCENT_COLOR
BUTTON_HOVER_BG_COLOR = ACCENT_HOVER

FONT_TITLE = ("Segoe UI", 26, "bold")
FONT_SECTION = ("Segoe UI", 16, "bold")
FONT_BUTTON = ("Segoe UI", 14, "bold")
FONT_FOOTER = ("Segoe UI", 12)

meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

años = [str(a) for a in range(datetime.now().year - 5, datetime.now().year + 6)]
