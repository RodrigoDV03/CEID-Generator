import os
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from .constants import *
from .generador import *

def iniciar_interfaz_fase_final(callback_volver=None):
    # --- VENTANA PRINCIPAL ---
    root = tk.Tk()
    root.title("Generador de Conformidades - FASE FINAL")
    root.geometry("600x500")
    root.configure(bg=BG_COLOR)
    root.resizable(False, False)

    # --- ENCABEZADO ---
    tk.Label(
        root, text="Generador de Conformidades - FASE FINAL",
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
        text="Seleccionar planilla del mes",
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

    # --- SELECCIÓN DE CARPETA DE SALIDA ---
    frame_salida = tk.Frame(root, bg=BG_COLOR)
    frame_salida.pack(fill="x", padx=30, pady=(15, 0))

    label_salida = tk.Label(
        frame_salida,
        text="Ninguna carpeta seleccionada",
        fg="red", bg=BG_COLOR,
        font=("Segoe UI", 10, "italic")
    )
    label_salida.pack(side="left", padx=(0, 10))

    def seleccionar_salida():
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de salida base")
        if ruta:
            label_salida.config(text=f"📂 {os.path.basename(ruta)}", fg="green")
            boton_generar.carpeta_salida = ruta

    tk.Button(
        frame_salida,
        text="Seleccionar carpeta",
        command=seleccionar_salida,
        font=FONT_BUTTON,
        bg=ACCENT_COLOR, fg="white",
        activebackground="#357ABD", activeforeground="white",
        relief="flat", padx=10, pady=4
    ).pack(side="right")

    # --- BOTÓN GENERAR ---
    def iniciar_generacion():
        hoja = hoja_var.get()
        ruta = getattr(boton_generar, "ruta_excel", None)
        carpeta = getattr(boton_generar, "carpeta_salida", None)
        mes = mes_var.get()
        año = año_var.get()
        if not hoja or not ruta or not carpeta or not mes or not año:
            messagebox.showerror("Error", "Debe seleccionar todos los campos: archivo, hoja, carpeta, mes y año.")
            return
        try:
            generados, errores = procesar_planilla(ruta, hoja, carpeta, mes, año)
            msg = f"Documentos generados: {len(generados)}"
            if errores:
                msg += "\n\nErrores:\n" + "\n".join([f"{d}: {e}" for d, e in errores])
            messagebox.showinfo("Proceso finalizado", msg)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

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
    boton_generar.pack(pady=30)

    def volver():
        root.destroy()
        if callback_volver:
            callback_volver()

    tk.Button(
        root, text="⬅ Volver al Menú Principal",
        command=volver,
        font=("Segoe UI", 10),
        bg="#cccccc", fg="#222",
        relief="flat", padx=8, pady=4
    ).pack(pady=(10, 15), side="bottom")

    # --- PIE DE PÁGINA ---
    tk.Label(
        root,
        text="CEID Generator - FASE FINAL",
        font=FONT_FOOTER,
        bg=BG_COLOR,
        fg="#a0a8b8"
    ).pack(side="bottom", pady=(0, 10))

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_fase_final()