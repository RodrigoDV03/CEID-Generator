import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import pandas as pd
from .generador_documentos import *

def iniciar_interfaz_fase_inicial(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title("Fase Inicial | Generador de Archivos CEID")
    ventana.geometry("650x700")
    ventana.resizable(False, False)

    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    ruta_excel = None
    carpeta_destino = None

    def titulo(texto):
        return ctk.CTkLabel(ventana, text=texto, font=ctk.CTkFont(size=22, weight="bold"))

    def etiqueta(texto, size=13):
        return ctk.CTkLabel(ventana, text=texto, font=ctk.CTkFont(size=size))

    # Título principal
    titulo("📄 Generador de Archivos - Fase Inicial").pack(pady=(25, 10))

    # Sección archivo Excel
    frame_excel = ctk.CTkFrame(ventana)
    frame_excel.pack(padx=30, pady=10, fill="x")

    etiqueta("Seleccionar archivo Excel del mes:", 15).pack(in_=frame_excel, anchor="w", padx=10, pady=(10, 2))
    label_excel = ctk.CTkLabel(frame_excel, text="📂 No seleccionado", text_color="gray")
    label_excel.pack(side="left", padx=10)

    def seleccionar_archivo():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_menu.configure(values=hojas)
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color="black")
                boton_generar.configure(state="normal")
                ruta_excel = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    ctk.CTkButton(frame_excel, text="Seleccionar archivo", command=seleccionar_archivo, width=160).pack(side="right", padx=10)

    # Hoja de trabajo
    frame_hoja = ctk.CTkFrame(ventana)
    frame_hoja.pack(padx=30, pady=(0, 10), fill="x")

    etiqueta("Hoja de trabajo:", 15).pack(in_=frame_hoja, anchor="w", padx=10, pady=(10, 2))
    hoja_menu = ctk.CTkOptionMenu(frame_hoja, variable=hoja_var, values=[])
    hoja_menu.pack(padx=10, pady=(0, 10), fill="x")

    # Mes y año
    frame_fecha = ctk.CTkFrame(ventana)
    frame_fecha.pack(padx=30, pady=10, fill="x")

    etiqueta("Mes:", 15).pack(in_=frame_fecha, anchor="w", padx=10, pady=(10, 2))
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    ctk.CTkOptionMenu(frame_fecha, variable=mes_var, values=meses).pack(padx=10, fill="x")

    etiqueta("Año:", 15).pack(in_=frame_fecha, anchor="w", padx=10, pady=(10, 2))
    años = [str(a) for a in range(datetime.now().year - 5, datetime.now().year + 6)]
    ctk.CTkOptionMenu(frame_fecha, variable=año_var, values=años).pack(padx=10, fill="x", pady=(0, 10))

    # Selección de carpeta destino
    frame_destino = ctk.CTkFrame(ventana)
    frame_destino.pack(padx=30, pady=10, fill="x")

    etiqueta("Seleccionar carpeta de destino:", 15).pack(in_=frame_destino, anchor="w", padx=10, pady=(10, 2))
    label_destino = ctk.CTkLabel(frame_destino, text="📂 No seleccionado", text_color="gray")
    label_destino.pack(side="left", padx=10)

    def seleccionar_carpeta():
        nonlocal carpeta_destino
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if carpeta:
            carpeta_destino = carpeta
            label_destino.configure(text=f"📁 {os.path.dirname(carpeta)}", text_color="black")

    ctk.CTkButton(frame_destino, text="Seleccionar carpeta", command=seleccionar_carpeta, width=160).pack(side="right", padx=10)

    # Botón generar documentos
    def generar():
        if not ruta_excel:
            return messagebox.showwarning("Falta archivo", "Por favor selecciona un archivo Excel.")
        if not hoja_var.get():
            return messagebox.showwarning("Falta hoja", "Selecciona una hoja del archivo.")
        if not carpeta_destino:
            return messagebox.showwarning("Falta carpeta", "Selecciona una carpeta de destino.")
        try:
            generar_documentos(ruta_excel, hoja_var.get(), carpeta_destino, mes_var.get(), año_var.get())
            messagebox.showinfo("Éxito", "¡Los documentos fueron generados correctamente!")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    boton_generar = ctk.CTkButton(
        ventana, text="🚀 Generar Documentos", command=generar,
        state="disabled", height=40, font=ctk.CTkFont(size=14, weight="bold")
    )
    boton_generar.pack(pady=25, padx=60, fill="x")

    def volver():
        ventana.destroy()
        if callback_volver:
            callback_volver()

    ctk.CTkButton(
        ventana, text="⬅ Volver al menú",
        command=volver, fg_color="#cccccc", text_color="#222"
    ).pack(pady=(5, 15))

    # Footer
    ctk.CTkLabel(ventana, text="CEID Generator - FASE INICIAL", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(0, 10))

    ventana.mainloop()
