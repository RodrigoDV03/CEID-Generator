import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import pandas as pd
from generador_documentos import *


# --- VENTANA PRINCIPAL ---
root = tk.Tk()
root.title("Generador de Archivos Fase Inicial CEID")
root.geometry("560x380")
root.configure(bg="#f4f6fa")
root.resizable(False, False)

# --- ESTILOS ---
PRIMARY_COLOR = "#2d415a"
ACCENT_COLOR = "#4a90e2"
BG_COLOR = "#f4f6fa"
TEXT_COLOR = "#333"
DISABLED_COLOR = "#bbb"

FONT_TITLE = ("Segoe UI", 17, "bold")
FONT_LABEL = ("Segoe UI", 11)
FONT_BUTTON = ("Segoe UI", 10)
FONT_FOOTER = ("Segoe UI", 9)

# --- ENCABEZADO ---
tk.Label(
    root, text="Generador de Archivos Fase Inicial CEID",
    font=FONT_TITLE, bg=BG_COLOR, fg=PRIMARY_COLOR
).pack(pady=(20, 10))

ttk.Separator(root, orient="horizontal").pack(fill="x", padx=30)

# --- SELECCIÓN DE ARCHIVO ---
frame_archivo = tk.Frame(root, bg=BG_COLOR)
frame_archivo.pack(fill="x", padx=30, pady=(20, 5))

label_archivo = tk.Label(
    frame_archivo,
    text="Ningún archivo seleccionado",
    fg="red", bg=BG_COLOR,
    font=("Segoe UI", 10, "italic")
)
label_archivo.pack(side="left", padx=(0, 10))

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if ruta:
        try:
            hojas = pd.ExcelFile(ruta).sheet_names
            hoja_var.set(hojas[0])
            hoja_menu['menu'].delete(0, 'end')
            for h in hojas:
                hoja_menu['menu'].add_command(label=h, command=lambda val=h: hoja_var.set(val))
            label_archivo.config(text=f"📁 {os.path.basename(ruta)}", fg="green")
            boton_generar.config(state="normal")
            boton_generar.ruta_excel = ruta
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

tk.Button(
    frame_archivo,
    text="Seleccionar archivo",
    command=seleccionar_archivo,
    font=FONT_BUTTON,
    bg=ACCENT_COLOR, fg="white",
    activebackground="#357ABD", activeforeground="white",
    relief="flat", padx=10, pady=4
).pack(side="right")

# --- SELECCIÓN DE HOJA ---
frame_hoja = tk.Frame(root, bg=BG_COLOR)
frame_hoja.pack(fill="x", padx=30, pady=(15, 0))

tk.Label(
    frame_hoja,
    text="Selecciona la hoja de trabajo:",
    font=FONT_LABEL,
    bg=BG_COLOR, fg=TEXT_COLOR
).pack(side="left")

hoja_var = tk.StringVar()
hoja_menu = tk.OptionMenu(frame_hoja, hoja_var, "")
hoja_menu.config(font=FONT_BUTTON, width=25)
hoja_menu.pack(side="left", padx=(10, 0))

# --- SELECCIÓN DE MES Y AÑO ---
frame_fecha = tk.Frame(root, bg=BG_COLOR)
frame_fecha.pack(fill="x", padx=30, pady=(15, 0))

tk.Label(frame_fecha, text="Mes:", font=FONT_LABEL, bg=BG_COLOR, fg=TEXT_COLOR).pack(side="left")

meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_var = tk.StringVar(value=meses[datetime.now().month - 1])
mes_menu = tk.OptionMenu(frame_fecha, mes_var, *meses)
mes_menu.config(font=FONT_BUTTON)
mes_menu.pack(side="left", padx=(5, 20))

tk.Label(frame_fecha, text="Año:", font=FONT_LABEL, bg=BG_COLOR, fg=TEXT_COLOR).pack(side="left")

año_actual = datetime.now().year
años = list(range(año_actual - 5, año_actual + 6))
año_var = tk.StringVar(value=str(año_actual))
año_menu = tk.OptionMenu(frame_fecha, año_var, *años)
año_menu.config(font=FONT_BUTTON)
año_menu.pack(side="left", padx=(5, 0))


# --- SELECCIÓN DE CARPETA DESTINO ---
frame_carpeta = tk.Frame(root, bg=BG_COLOR)
frame_carpeta.pack(fill="x", padx=30, pady=(10, 0))

label_carpeta = tk.Label(
    frame_carpeta,
    text="Ninguna carpeta seleccionada",
    fg="red", bg=BG_COLOR,
    font=("Segoe UI", 10, "italic")
)
label_carpeta.pack(side="left", padx=(0, 10))

carpeta_destino = {"ruta": None}

def seleccionar_carpeta():
    ruta = filedialog.askdirectory(
        title="Seleccionar carpeta de destino"
    )
    if ruta:
        carpeta_destino["ruta"] = ruta
        label_carpeta.config(text=f"📂 {ruta}", fg="green")

tk.Button(
    frame_carpeta,
    text="Seleccionar carpeta de destino",
    command=seleccionar_carpeta,
    font=FONT_BUTTON,
    bg=ACCENT_COLOR, fg="white",
    activebackground="#357ABD", activeforeground="white",
    relief="flat", padx=10, pady=4
).pack(side="right")


# --- BOTÓN GENERAR ---
def iniciar_generacion():
    hoja = hoja_var.get()
    ruta = getattr(boton_generar, "ruta_excel", None)
    carpeta = carpeta_destino.get("ruta")
    mes = mes_var.get()
    año = año_var.get()
    if not hoja or not ruta or not carpeta or not mes or not año:
        messagebox.showerror("Error", "Debe seleccionar todos los campos: archivo, hoja, carpeta, mes y año.")
        return
    
    generar_documentos(ruta, hoja, carpeta, mes, año)

boton_generar = tk.Button(
    root,
    text="📄 Generar documentos",
    state="disabled",
    command=iniciar_generacion,
    font=("Segoe UI", 12, "bold"),
    bg=PRIMARY_COLOR, fg="white",
    activebackground="#466a8f",
    activeforeground="white",
    relief="flat",
    padx=12, pady=8
)
boton_generar.pack(pady=35)

# --- PIE DE PÁGINA ---
tk.Label(
    root,
    text="CEID Generator - v1.0",
    font=FONT_FOOTER,
    bg=BG_COLOR,
    fg="#a0a8b8"
).pack(side="bottom", pady=(0, 10))

root.mainloop()