from ui.layout import AppLayout
from ui.components import card_opcion
import customtkinter as ctk

from ui.views.correos_view import mostrar_correos
from ui.views.planilla_view import mostrar_planilla
from ui.views.fase_inicial_view import mostrar_fase_inicial
from ui.views.fase_final_view import mostrar_fase_final
from ui.views.control_pagos_view import mostrar_control_pagos

def iniciar_interfaz_general():

    app = AppLayout("CEID - Sistema")

    # --------- NAVEGACIÓN ---------
    app.agregar_opcion_sidebar("🏠 Inicio", lambda: mostrar_inicio(app))
    app.agregar_opcion_sidebar("📊 Planilla", lambda: mostrar_planilla(app))
    app.agregar_opcion_sidebar("📄 Fase Inicial", lambda: mostrar_fase_inicial(app))
    app.agregar_opcion_sidebar("✅ Fase Final", lambda: mostrar_fase_final(app))
    app.agregar_opcion_sidebar("💰 Control de Pagos", lambda: mostrar_control_pagos(app))
    app.agregar_opcion_sidebar("✉️ Correos", lambda: mostrar_correos(app))

    mostrar_inicio(app)

    app.ejecutar()


# ---------------- VISTAS ----------------

def mostrar_inicio(app):
    app.limpiar_contenido()
    app.titulo_header.configure(text="Panel principal")

    frame = app.main

    fila = ctk.CTkFrame(frame, fg_color="transparent")
    fila.pack(expand=True, fill="both")

    card_opcion(
        fila,
        "Generar Planilla",
        "Crear planillas desde Excel",
        lambda: mostrar_planilla(app)
    )

    card_opcion(
        fila,
        "Fase Inicial",
        "Documentos iniciales",
        lambda: mostrar_fase_inicial(app)
    )

    card_opcion(
        fila,
        "Fase Final",
        "Documentos finales",
        lambda: mostrar_fase_final(app)
    )

    card_opcion(
        fila,
        "Control de Pagos",
        "Actualiza armadas y saldos",
        lambda: mostrar_control_pagos(app)
    )

    card_opcion(
        fila,
        "Correos",
        "Enviar correos masivos",
        lambda: mostrar_correos(app)
    )