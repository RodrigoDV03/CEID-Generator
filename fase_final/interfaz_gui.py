import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from .generador import *

def iniciar_interfaz_fase_final(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    # --- VENTANA PRINCIPAL ---
    root = ctk.CTk()
    root.title("Fase Final | Generador de Archivos CEID")
    root.geometry("650x800")
    root.resizable(False, False)

    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    ruta_excel = None
    archivo_docente = None
    carpeta_destino = None

    def titulo(texto):
        return ctk.CTkLabel(root, text=texto, font=ctk.CTkFont(size=22, weight="bold"))
    
    def etiqueta(texto, size=13):
        return ctk.CTkLabel(root, text=texto, font=ctk.CTkFont(size=size))


    # Título principal
    titulo("📄 Generador de Archivos - Fase Final").pack(pady=(25, 10))

    # Sección archivo Excel
    frame_excel = ctk.CTkFrame(root)
    frame_excel.pack(padx=30, pady=10, fill="x")

    etiqueta("Seleccionar archivo Excel del mes:", 15).pack(in_=frame_excel, anchor="w", padx=10, pady=(10, 2))
    label_archivo = ctk.CTkLabel(frame_excel, text="📂 No seleccionado", text_color="gray")
    label_archivo.pack(side="left", padx=10)

    def seleccionar_archivo():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_menu.configure(values=hojas)
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_archivo.configure(text=f"📁 {os.path.basename(ruta)}", text_color="black")
                boton_generar.configure(state="normal")
                ruta_excel = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    ctk.CTkButton(frame_excel, text="Seleccionar archivo", command=seleccionar_archivo, width=160).pack(side="right", padx=10) 

    # Hoja de trabajo
    frame_hoja = ctk.CTkFrame(root)
    frame_hoja.pack(padx=30, pady=(0, 10), fill="x")

    etiqueta("Hoja de trabajo:", 15).pack(in_=frame_hoja, anchor="w", padx=10, pady=(10, 2))
    hoja_menu = ctk.CTkOptionMenu(frame_hoja, variable=hoja_var, values=[])
    hoja_menu.pack(padx=10, pady=(0, 10), fill="x")

    # --- SELECCIÓN EXCEL DE DOCENTE DE CONTRATO PARA GENERACIÓN DE CONTROL DE PAGOS ---

    frame_docente = ctk.CTkFrame(root)
    frame_docente.pack(padx=30, pady=10, fill="x")

    etiqueta("Seleccionar excel de docentes de contrato:", 15).pack(in_=frame_docente, anchor="w", padx=10, pady=(10, 2))
    label_docente = ctk.CTkLabel(frame_docente, text="📂 No seleccionado", text_color="gray")
    label_docente.pack(side="left", padx=10)

    def seleccionar_docente():
        nonlocal archivo_docente
        archivo_docente = filedialog.askopenfilename(
            title="Seleccionar archivo de docentes",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if archivo_docente:
            try:
                pd.ExcelFile(archivo_docente)
                label_docente.configure(text=f"📁 {os.path.basename(archivo_docente)}", text_color="black")
                boton_generar.archivo_docente = archivo_docente
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    ctk.CTkButton(frame_docente, text="Seleccionar archivo", command=seleccionar_docente, width=160).pack(side="right", padx=10)

    # Mes y año
    frame_fecha = ctk.CTkFrame(root)
    frame_fecha.pack(padx=30, pady=10, fill="x")

    etiqueta("Mes:", 15).pack(in_=frame_fecha, anchor="w", padx=10, pady=(10, 2))
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    ctk.CTkOptionMenu(frame_fecha, variable=mes_var, values=meses).pack(padx=10, fill="x")

    etiqueta("Año:", 15).pack(in_=frame_fecha, anchor="w", padx=10, pady=(10, 2))
    años = [str(a) for a in range(datetime.now().year - 5, datetime.now().year + 6)]
    ctk.CTkOptionMenu(frame_fecha, variable=año_var, values=años).pack(padx=10, fill="x", pady=(0, 10))

    # --- SELECCIÓN DE CARPETA DE SALIDA ---
    frame_salida = ctk.CTkFrame(root)
    frame_salida.pack(padx=30, pady=10, fill="x")

    etiqueta("Seleccionar carpeta de destino:", 15).pack(in_=frame_salida, anchor="w", padx=10, pady=(10, 2))
    label_destino = ctk.CTkLabel(frame_salida, text="📂 No seleccionado", text_color="gray")
    label_destino.pack(side="left", padx=10)

    def seleccionar_salida():
        nonlocal carpeta_destino
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if ruta:
            carpeta_destino = ruta
            label_destino.configure(text=f"📂 {os.path.dirname(ruta)}", text_color="black")

    ctk.CTkButton(frame_salida, text="Seleccionar carpeta", command=seleccionar_salida, width=160).pack(side="right", padx=10)

    # --- BOTÓN GENERAR ---
    def iniciar_generacion():
        # if not ruta_excel or not archivo_docente or not hoja_var.get() or not carpeta_destino:
        #     messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
        #     return
        try:
            procesar_planilla(ruta_excel, archivo_docente, hoja_var.get(), carpeta_destino, mes_var.get(), año_var.get())
            messagebox.showinfo("Éxito", f"Documentos de fase final generados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

    boton_generar = ctk.CTkButton(
        root, text="🚀 Generar Documentos", command=iniciar_generacion,
        state="disabled", height=40, font=ctk.CTkFont(size=14, weight="bold")
    )
    boton_generar.pack(pady=25, padx=60, fill="x")

    def volver():
        root.destroy()
        if callback_volver:
            callback_volver()

    ctk.CTkButton(
        root, text="⬅ Volver al menú",
        command=volver, fg_color="#cccccc", text_color="#222"
    ).pack(pady=(5, 15))

    # Footer
    ctk.CTkLabel(root, text="CEID Generator - FASE FINAL", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(0, 10))

    root.mainloop()