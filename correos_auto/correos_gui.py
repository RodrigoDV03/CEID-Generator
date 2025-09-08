import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog, messagebox
from datetime import datetime
from .envio_correos import procesar_correos_administrativos, procesar_correos_docente, enviar_correo_administrativo, enviar_correo_docente
from utils.gui_constants import *

def iniciar_interfaz_correos(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    data_para_envio = []

    root = ctk.CTk()
    root.title("Envío de correos | CEID")
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    tipo_var = ctk.StringVar(value="Docente")

    ruta_excel = None
    pdfs_seleccionados = []

    # --- TÍTULO PRINCIPAL ---
    titulo(root, "Envío de correos | CEID")

    # --- CARD PRINCIPAL ---
    card = ctk.CTkFrame(root, fg_color=SECTION_COLOR, corner_radius=15)
    card.pack(padx=40, pady=20, fill="both", expand=False)

    # --- TIPO DE DESTINATARIO ---
    etiqueta(card, "Tipo de destinatario:")
    crear_option_menu(card, tipo_var, ["Docente", "Administrativo"])

    # --- ARCHIVO EXCEL ---
    etiqueta(card, "Seleccionar excel de docentes o administrativos:")

    def seleccionar_archivo():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(title="Seleccionar excel de docentes o administrativos", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_var.set("list")
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color=WHITE_COLOR)
                boton_gen_data.configure(state="normal")
                ruta_excel = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    boton_excel, label_excel = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_archivo)

    # --- ARCHIVOS PDF ---
    etiqueta(card, "Seleccionar PDFs de órdenes de servicio:")

    def seleccionar_pdfs():
        nonlocal pdfs_seleccionados
        rutas = filedialog.askopenfilenames(
            title="Seleccionar PDFs de órdenes de servicio",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if rutas:
            pdfs_seleccionados = list(rutas)
            label_pdfs.configure(text=f"📁 {len(pdfs_seleccionados)} PDFs seleccionados", text_color=WHITE_COLOR)

    boton_pdfs, label_pdfs = crear_boton_archivo(card, "📂 Ningún PDF seleccionado", seleccionar_pdfs)

    # --- FECHA ---
    etiqueta(card, "Mes:")
    crear_option_menu(card, mes_var, meses)

    etiqueta(card, "Año:")
    crear_option_menu(card, año_var, años)

    # --- BOTÓN GENERAR DATA ---
    def generar_data():
        nonlocal data_para_envio
        if not ruta_excel or not hoja_var.get() or not pdfs_seleccionados:
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return
        def tarea():
            nonlocal data_para_envio
            try:
                if tipo_var.get() == "Docente":
                    data_para_envio = procesar_correos_docente(ruta_excel, hoja_var.get(), pdfs_seleccionados)
                else:
                    data_para_envio = procesar_correos_administrativos(ruta_excel, hoja_var.get(), pdfs_seleccionados)

                if not data_para_envio:
                    messagebox.showwarning("Aviso", "No se generó ninguna coincidencia para envío.")
                else:
                    messagebox.showinfo("Éxito", f"Datos generados. Revisar antes de enviar.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar: {e}")

        threading.Thread(target=tarea).start()

    boton_gen_data = boton_generador(card, "Generar Datos", generar_data)

    # --- CONSOLA EMBEBIDA ---
    consola_frame = ctk.CTkFrame(root, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(padx=40, pady=(10, 20), fill="both", expand=True)
    consola_text = ctk.CTkTextbox(consola_frame, height=120, wrap="word", fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola_text.pack(padx=10, pady=10, fill="both", expand=True)
    consola_text.configure(state="disabled")

    class TextRedirector:
        def __init__(self, text_widget):
            self.text_widget = text_widget
        def write(self, message):
            if self.text_widget.winfo_exists():
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", message)
                self.text_widget.see("end")
                self.text_widget.configure(state="disabled")
        def flush(self): pass

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)

    # --- BOTÓN ENVIAR CORREOS ---
    def enviar():
        if not data_para_envio:
            messagebox.showerror("Error", "No hay datos para enviar. Primero genere la data.")
            return

        respuesta = messagebox.askyesno("Confirmación", f"Se enviarán {len(data_para_envio)} correos. ¿Continuar?")
        if not respuesta:
            return

        boton_envio.configure(state="disabled")  # Deshabilita el botón mientras se envía

        def tarea_envio():
            try:
                for item in data_para_envio:
                    if tipo_var.get() == "Docente":
                        enviar_correo_docente(
                            nombre=item['nombre'],
                            pdf_path=item['pdf_path'],
                            destinatario=item['correo'],
                            mes=mes_var.get(),
                            anio=año_var.get(),
                            servicio=item['servicio']
                        )
                    else:
                        enviar_correo_administrativo(
                            nombre=item['nombre'],
                            pdf_path=item['pdf_path'],
                            destinatario=item['correo'],
                            mes=mes_var.get(),
                            anio=año_var.get(),
                        )

                messagebox.showinfo("Éxito", "Todos los correos fueron enviados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Fallo en el envío: {e}")
            finally:
                boton_envio.configure(state="normal")

        threading.Thread(target=tarea_envio).start()

    boton_envio = boton_generador(card, "Enviar Correos", enviar)

    # --- BOTÓN VOLVER ---
    boton_volver(root, callback_volver)

    # --- FOOTER ---
    footer(root)

    root.mainloop()