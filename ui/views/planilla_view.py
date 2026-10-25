import customtkinter as ctk
from tkinter import filedialog
import os
import threading

from core.planillas.generador_planilla import generar_planilla
from ui.progress import BarraProgreso
from utils.gui_constants import *
from utils import custom_modals as messagebox


def mostrar_planilla(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Generador de Planillas")

    # =========================
    # VARIABLES DE ESTADO
    # =========================
    archivo_cursos = ""
    archivo_docentes = ""
    archivo_clasif = ""
    archivo_coordinacion = ""
    planillas_anteriores = []
    carpeta_destino = ""

    mes_var = ctk.StringVar(value=mes_actual())
    bono_var = ctk.StringVar(value="0")

    root = app.root if hasattr(app, "root") else app.winfo_toplevel()

    # =========================
    # CONTENEDOR PRINCIPAL
    # =========================
    contenedor = ctk.CTkScrollableFrame(app.main, fg_color="transparent")
    contenedor.pack(fill="both", expand=True, padx=30, pady=20)

    # =========================
    # FUNCIONES
    # =========================
    def estado(label, texto, color=TEXT_LIGHT):
        label.configure(text=texto, text_color=color)

    def validar_formulario():
        valido = archivo_cursos and archivo_docentes
        btn_generar.configure(state="normal" if valido else "disabled")

    # =====================================================
    # ① CONFIGURACIÓN
    # =====================================================
    frame_conf = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_conf.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_conf, text="Configuración", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_conf, text="Mes", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_conf, variable=mes_var, values=meses)\
        .pack(padx=20, pady=(0, 10), anchor="w")

    # Bono solo enero
    frame_bono = ctk.CTkFrame(frame_conf, fg_color="transparent")

    ctk.CTkLabel(frame_bono, text="Monto bono (solo enero)", text_color=TEXT_COLOR)\
        .pack(anchor="w")

    entry_bono = ctk.CTkEntry(frame_bono, textvariable=bono_var, width=200)
    entry_bono.pack(anchor="w")

    def toggle_bono(*_):
        if mes_var.get() == "Enero":
            frame_bono.pack(anchor="w", padx=20, pady=10)
        else:
            frame_bono.pack_forget()
            bono_var.set("0")

    mes_var.trace_add("write", toggle_bono)
    toggle_bono()

    # =====================================================
    # ② ARCHIVOS
    # =====================================================
    frame_archivos = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_archivos.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_archivos, text="Archivos obligatorios", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    def selector_archivo(master, texto, extensiones, setter):
        frame = ctk.CTkFrame(master, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=5)

        ctk.CTkLabel(frame, text=texto, text_color=TEXT_COLOR)\
            .pack(anchor="w")

        estado_lbl = ctk.CTkLabel(frame, text="No seleccionado", text_color=TEXT_LIGHT)
        estado_lbl.pack(side="left", padx=10)

        def seleccionar():
            archivo = filedialog.askopenfilename(filetypes=[extensiones])
            if archivo:
                setter(archivo)
                estado(estado_lbl, os.path.basename(archivo), "#4CAF50")
                validar_formulario()

        ctk.CTkButton(
            frame,
            text="Seleccionar",
            width=140,
            fg_color=PRIMARY_COLOR,
            hover_color=ACCENT_COLOR,
            command=seleccionar
        ).pack(side="right")

    def set_cursos(v):
        nonlocal archivo_cursos
        archivo_cursos = v

    def set_docentes(v):
        nonlocal archivo_docentes
        archivo_docentes = v

    def set_clasif(v):
        nonlocal archivo_clasif
        archivo_clasif = v

    selector_archivo(frame_archivos, "Cursos (CSV)", ("CSV", "*.csv"), set_cursos)
    selector_archivo(frame_archivos, "Docentes (Excel)", ("Excel", "*.xlsx *.xls"), set_docentes)

    # =====================================================
    # ③ OPCIONAL
    # =====================================================
    frame_extra = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_extra.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_extra, text="Archivos adicionales (Opcionales)", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    def set_coord(v):
        nonlocal archivo_coordinacion
        archivo_coordinacion = v

    selector_archivo(frame_extra, "Examen de Clasificación", ("Excel", "*.xlsx *.xls"), set_clasif)
    selector_archivo(frame_extra, "Apoyo Docente", ("Excel", "*.xlsx *.xls"), set_coord)

    # Planillas anteriores del mes: si se eligen, esta es la 2ª (3ª...) planilla
    # y solo incluye cursos nuevos o con cambio de docente.
    frame_anteriores = ctk.CTkFrame(frame_extra, fg_color="transparent")
    frame_anteriores.pack(fill="x", padx=20, pady=(5, 15))

    ctk.CTkLabel(frame_anteriores, text="CSV de planillas anteriores del mes (para la 2ª planilla)", text_color=TEXT_COLOR)\
        .pack(anchor="w")

    lbl_anteriores = ctk.CTkLabel(frame_anteriores, text="Ninguno (1ª planilla)", text_color=TEXT_LIGHT)
    lbl_anteriores.pack(side="left", padx=10)

    def seleccionar_anteriores():
        nonlocal planillas_anteriores
        rutas = filedialog.askopenfilenames(filetypes=[("CSV", "*.csv")])
        if rutas:
            planillas_anteriores = list(rutas)
            texto = os.path.basename(rutas[0]) if len(rutas) == 1 else f"{len(rutas)} archivos"
            estado(lbl_anteriores, f"{texto} → se generará la planilla N.° {len(rutas) + 1}", "#4CAF50")

    def quitar_anteriores():
        nonlocal planillas_anteriores
        planillas_anteriores = []
        estado(lbl_anteriores, "Ninguno (1ª planilla)")

    ctk.CTkButton(
        frame_anteriores,
        text="Quitar",
        width=80,
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=quitar_anteriores
    ).pack(side="right", padx=(10, 0))

    ctk.CTkButton(
        frame_anteriores,
        text="Seleccionar",
        width=140,
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_anteriores
    ).pack(side="right")

    # =====================================================
    # ④ DESTINO
    # =====================================================
    frame_destino = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_destino.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_destino, text="Destino", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    lbl_destino = ctk.CTkLabel(frame_destino, text="Automático", text_color=TEXT_LIGHT)
    lbl_destino.pack(anchor="w", padx=20)

    def elegir_destino():
        nonlocal carpeta_destino
        carpeta = filedialog.askdirectory()
        if carpeta:
            carpeta_destino = carpeta
            estado(lbl_destino, carpeta, "#4CAF50")

    ctk.CTkButton(
        frame_destino,
        text="Elegir carpeta",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=elegir_destino
    ).pack(anchor="w", padx=20, pady=10)

    # =====================================================
    # BOTÓN GENERAR
    # =====================================================
    def mostrar_resultado(resultado):
        btn_generar.configure(state="normal")

        mensaje = resultado.mensaje
        if resultado.comparacion is not None:
            c = resultado.comparacion
            mensaje += (
                f"\n\nComparación con {c.planillas_anteriores} planilla(s) anterior(es):\n"
                f"- Cursos ya emitidos (excluidos): {c.excluidos}\n"
                f"- Cursos con cambio de docente (incluidos): {c.cambio_docente}\n"
                f"- Cursos nuevos (incluidos): {c.nuevos}\n"
                "El detalle está en la hoja \"Comparación\"."
            )

        if not resultado.ok:
            barra.completar("Error")
            messagebox.showerror("Error", mensaje)
        elif resultado.sin_coincidencia:
            barra.completar("Completado con advertencias")
            lista = "\n".join(f"- {nombre}" for nombre in resultado.sin_coincidencia)
            messagebox.showwarning(
                "Planilla generada con advertencias",
                f"{mensaje}\n\n"
                f"{len(resultado.sin_coincidencia)} docente(s) del archivo de cursos NO se pudieron "
                f"emparejar con la base de docentes y quedaron FUERA del cálculo de pago:\n{lista}\n\n"
                "Revisa que el nombre esté escrito igual en ambos archivos."
            )
        else:
            barra.completar("Completado")
            messagebox.showinfo("Éxito", mensaje)

    def procesar():
        if not messagebox.askyesno("Confirmar", "¿Generar la planilla?"):
            return

        try:
            monto_bono = float(bono_var.get() or 0)
        except ValueError:
            messagebox.showerror("Error", "El monto del bono debe ser un número.")
            return

        # Leer los valores de la pantalla aquí, en el hilo principal: los
        # StringVar de Tkinter no deben leerse desde el hilo de fondo.
        mes = mes_var.get()
        anteriores = list(planillas_anteriores)

        btn_generar.configure(state="disabled")
        barra.iniciar()

        def al_error(e):
            btn_generar.configure(state="normal")
            barra.completar("Error")
            messagebox.showerror("Error", str(e))

        def tarea():
            try:
                resultado = generar_planilla(
                    archivo_cursos,
                    archivo_docentes,
                    archivo_clasif,
                    archivo_coordinacion,
                    mes,
                    monto_bono,
                    carpeta_destino or None,
                    planillas_anteriores=anteriores,
                    on_progreso=barra.actualizar,
                )
                root.after(0, lambda: mostrar_resultado(resultado))
            except Exception as e:
                root.after(0, lambda: al_error(e))

        threading.Thread(target=tarea, daemon=True).start()

    btn_generar = ctk.CTkButton(
        contenedor,
        text="Generar Planilla",
        height=50,
        fg_color=ACCENT_COLOR,
        hover_color=PRIMARY_COLOR,
        state="disabled",
        command=procesar
    )
    btn_generar.pack(pady=30)

    # =====================================================
    # PROGRESO
    # =====================================================
    barra = BarraProgreso(contenedor, root)
    barra.frame.pack(fill="x", padx=20, pady=(0, 10))
