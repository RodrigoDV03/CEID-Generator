import copy
import os
import re
import platform
from html import escape, unescape

import customtkinter as ctk
import pypdfium2 as pdfium
from PIL import ImageTk
from tkinter import Canvas

from utils.gui_constants import *
from utils import custom_modals as messagebox


class PreviewCorreosModal:
    def __init__(self, parent, previsualizaciones, on_save=None, on_send=None):
        self.parent = parent
        self.data = copy.deepcopy(previsualizaciones)
        self.on_save = on_save
        self.on_send = on_send

        self.idx_correo = 0
        self.idx_adjunto = 0
        self.pdf_paths_actuales = []
        self.idx_pagina_pdf = 0
        self.total_paginas_pdf = 1
        self._img_ref = None
        self._render_job = None
        self._form_loaded = False
        self.zoom_factor = 2.4
        self._platform = platform.system()
        self._pdf_image_id = None

        for item in self.data:
            cuerpo_texto = self._strip_html(item.get("cuerpo_html", ""))
            item.setdefault("cuerpo_texto", cuerpo_texto)
            item.setdefault("cuerpo_texto_original", cuerpo_texto)
            item.setdefault("cuerpo_html_original", item.get("cuerpo_html", ""))

        self.window = ctk.CTkToplevel(parent)
        self.window.title("Previsualizar correos")
        self.window.geometry("1280x760")
        self.window.minsize(1100, 680)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=BG_COLOR)

        self._crear_ui()
        if self.data:
            self._mostrar_correo(0)

    def _crear_ui(self):
        header = ctk.CTkFrame(self.window, fg_color=CARD_COLOR)
        header.pack(fill="x", padx=12, pady=(12, 8))

        nav = ctk.CTkFrame(header, fg_color="transparent")
        nav.pack(fill="x", padx=10, pady=(10, 6))

        self.btn_prev_correo = ctk.CTkButton(nav, text="← Anterior", width=110, fg_color=PRIMARY_COLOR, command=self._anterior_correo)
        self.btn_prev_correo.pack(side="left")

        self.lbl_correo = ctk.CTkLabel(nav, text="Correo 0 de 0", text_color=TEXT_COLOR, font=("Segoe UI", 13, "bold"))
        self.lbl_correo.pack(side="left", expand=True)

        self.btn_next_correo = ctk.CTkButton(nav, text="Siguiente →", width=110, fg_color=PRIMARY_COLOR, command=self._siguiente_correo)
        self.btn_next_correo.pack(side="right")

        self.lbl_dest = ctk.CTkLabel(header, text="", text_color=TEXT_LIGHT, font=("Segoe UI", 11))
        self.lbl_dest.pack(anchor="w", padx=14, pady=(0, 10))

        split = ctk.CTkFrame(self.window, fg_color="transparent")
        split.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)

        panel_pdf = ctk.CTkFrame(split, fg_color=CARD_COLOR, corner_radius=10)
        panel_pdf.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        ctk.CTkLabel(panel_pdf, text="PDF adjunto", text_color=TEXT_COLOR, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        nav_adjuntos = ctk.CTkFrame(panel_pdf, fg_color="transparent")
        nav_adjuntos.pack(fill="x", padx=10, pady=(0, 6))

        self.btn_prev_adjunto = ctk.CTkButton(
            nav_adjuntos,
            text="Adjunto ◄",
            width=90,
            fg_color=PRIMARY_COLOR,
            command=self._anterior_adjunto,
        )
        self.btn_prev_adjunto.pack(side="left")

        self.lbl_adjunto = ctk.CTkLabel(nav_adjuntos, text="Adjunto 1 de 1", text_color=TEXT_LIGHT)
        self.lbl_adjunto.pack(side="left", expand=True)

        self.btn_next_adjunto = ctk.CTkButton(
            nav_adjuntos,
            text="► Adjunto",
            width=90,
            fg_color=PRIMARY_COLOR,
            command=self._siguiente_adjunto,
        )
        self.btn_next_adjunto.pack(side="right")

        canvas_wrap = ctk.CTkFrame(panel_pdf, fg_color="transparent")
        canvas_wrap.pack(fill="both", expand=True, padx=10, pady=4)
        canvas_wrap.grid_rowconfigure(0, weight=1)
        canvas_wrap.grid_columnconfigure(0, weight=1)

        self.canvas_pdf = Canvas(canvas_wrap, bg="#FFFFFF", highlightthickness=0)
        self.canvas_pdf.grid(row=0, column=0, sticky="nsew")

        self.scroll_y = ctk.CTkScrollbar(canvas_wrap, orientation="vertical", command=self.canvas_pdf.yview)
        self.scroll_y.grid(row=0, column=1, sticky="ns", padx=(6, 0))

        self.scroll_x = ctk.CTkScrollbar(canvas_wrap, orientation="horizontal", command=self.canvas_pdf.xview)
        self.scroll_x.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        self.canvas_pdf.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)
        self.canvas_pdf.bind("<Configure>", self._on_pdf_resize)
        self.canvas_pdf.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas_pdf.bind("<Button-4>", self._on_mouse_wheel)
        self.canvas_pdf.bind("<Button-5>", self._on_mouse_wheel)
        self.canvas_pdf.bind("<ButtonPress-1>", self._on_pan_start)
        self.canvas_pdf.bind("<B1-Motion>", self._on_pan_drag)

        nav_pdf = ctk.CTkFrame(panel_pdf, fg_color="transparent")
        nav_pdf.pack(fill="x", padx=10, pady=(2, 10))

        self.btn_prev_pagina = ctk.CTkButton(nav_pdf, text="◄", width=42, fg_color=PRIMARY_COLOR, command=self._anterior_pagina)
        self.btn_prev_pagina.pack(side="left")

        self.lbl_pagina = ctk.CTkLabel(nav_pdf, text="Página 1 de 1", text_color=TEXT_LIGHT)
        self.lbl_pagina.pack(side="left", expand=True)

        self.btn_next_pagina = ctk.CTkButton(nav_pdf, text="►", width=42, fg_color=PRIMARY_COLOR, command=self._siguiente_pagina)
        self.btn_next_pagina.pack(side="right")

        zoom_frame = ctk.CTkFrame(panel_pdf, fg_color="transparent")
        zoom_frame.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(zoom_frame, text="Zoom", text_color=TEXT_COLOR).pack(side="left")
        ctk.CTkButton(zoom_frame, text="-", width=32, fg_color=PRIMARY_COLOR, command=lambda: self._cambiar_zoom(-0.1)).pack(side="left", padx=(8, 4))
        ctk.CTkButton(zoom_frame, text="+", width=32, fg_color=PRIMARY_COLOR, command=lambda: self._cambiar_zoom(0.1)).pack(side="left", padx=(0, 8))

        self.zoom_slider = ctk.CTkSlider(zoom_frame, from_=0.5, to=3.0, number_of_steps=25, command=self._on_zoom_slider)
        self.zoom_slider.set(2.4)
        self.zoom_slider.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.lbl_zoom = ctk.CTkLabel(zoom_frame, text="240%", text_color=TEXT_LIGHT, width=56)
        self.lbl_zoom.pack(side="left")

        ctk.CTkButton(zoom_frame, text="Ajustar", width=72, fg_color=PRIMARY_COLOR, command=self._reset_zoom).pack(side="left", padx=(8, 0))

        panel_msg = ctk.CTkFrame(split, fg_color=CARD_COLOR, corner_radius=10)
        panel_msg.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        ctk.CTkLabel(panel_msg, text="Mensaje del correo", text_color=TEXT_COLOR, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        ctk.CTkLabel(panel_msg, text="Asunto", text_color=TEXT_COLOR).pack(anchor="w", padx=12)
        self.entry_asunto = ctk.CTkEntry(panel_msg, fg_color="#FFFFFF", text_color="#000000", border_color=PRIMARY_COLOR)
        self.entry_asunto.pack(fill="x", padx=12, pady=(4, 10))

        ctk.CTkLabel(panel_msg, text="Cuerpo (editable)", text_color=TEXT_COLOR).pack(anchor="w", padx=12)
        self.txt_cuerpo = ctk.CTkTextbox(panel_msg, fg_color="#FFFFFF", text_color="#000000", wrap="word", font=("Segoe UI", 11))
        self.txt_cuerpo.pack(fill="both", expand=True, padx=12, pady=(4, 10))
        # CTkTextbox blocks the 'font' option at the wrapper level, so configure
        # the underlying tk.Text widget directly when possible.
        if hasattr(self.txt_cuerpo, "_textbox"):
            self.txt_cuerpo._textbox.tag_config("resaltado", foreground="#073763", font=("Segoe UI", 10, "bold"))
        else:
            self.txt_cuerpo.tag_config("resaltado", foreground="#073763")

        self.logs_frame = ctk.CTkFrame(self.window, fg_color=CARD_COLOR)
        self.logs_visible = False

        logs_header = ctk.CTkFrame(self.logs_frame, fg_color="transparent")
        logs_header.pack(fill="x", padx=10, pady=(8, 4))

        ctk.CTkLabel(
            logs_header,
            text="Registro en vivo de envío",
            text_color=TEXT_COLOR,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")

        self.lbl_estado_envio = ctk.CTkLabel(logs_header, text="", text_color=TEXT_LIGHT)
        self.lbl_estado_envio.pack(side="right")

        self.txt_logs = ctk.CTkTextbox(
            self.logs_frame,
            fg_color=CONSOLE_BG,
            text_color=WHITE_COLOR,
            wrap="word",
            height=170,
            font=("Consolas", 11),
        )
        self.txt_logs.pack(fill="both", expand=False, padx=10, pady=(0, 10))
        self.txt_logs.configure(state="disabled")

        footer = ctk.CTkFrame(self.window, fg_color=CARD_COLOR)
        footer.pack(fill="x", padx=12, pady=(0, 12))

        btns = ctk.CTkFrame(footer, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(btns, text="Cerrar", fg_color=BG_COLOR, text_color=TEXT_COLOR, command=self._cerrar).pack(side="left")
        ctk.CTkButton(btns, text="Guardar cambios", fg_color=PRIMARY_COLOR, hover_color=ACCENT_COLOR, command=self._guardar).pack(side="right", padx=(8, 0))
        self.btn_enviar = ctk.CTkButton(
            btns,
            text="Enviar correos",
            fg_color=ACCENT_COLOR,
            hover_color=PRIMARY_COLOR,
            command=self._enviar,
        )
        self.btn_enviar.pack(side="right")

    def _mostrar_panel_logs(self):
        if not self.logs_visible:
            self.logs_frame.pack(fill="x", padx=12, pady=(0, 8), before=self.window.pack_slaves()[-1])
            self.logs_visible = True

    def _append_log(self, texto):
        self._mostrar_panel_logs()
        self.txt_logs.configure(state="normal")
        self.txt_logs.insert("end", texto)
        self.txt_logs.see("end")
        self.txt_logs.configure(state="disabled")

    def _limpiar_logs(self):
        self._mostrar_panel_logs()
        self.txt_logs.configure(state="normal")
        self.txt_logs.delete("1.0", "end")
        self.txt_logs.configure(state="disabled")

    def _finalizar_envio(self, resumen, error):
        self.btn_enviar.configure(state="normal")

        if error:
            self.lbl_estado_envio.configure(text="Error", text_color="#ff6b6b")
            self._append_log(f"\n❌ Error: {error}\n")
            messagebox.showerror("Error", error, parent=self.window)
            return

        self.lbl_estado_envio.configure(text="Completado", text_color="#4CAF50")
        self._append_log(
            f"\n✅ Envío finalizado: {resumen['exitosos']} exitosos, {resumen['fallidos']} fallidos\n"
        )

        for item in resumen.get("resultados", []):
            if item.get("success"):
                thread_id = item.get("thread_id") or "N/A"
                message_id = item.get("message_id") or "N/A"
                self._append_log(
                    f"• {item.get('nombre', item.get('destinatario', 'Destinatario'))}: "
                    f"threadId={thread_id} messageId={message_id}\n"
                )
            else:
                self._append_log(
                    f"• {item.get('nombre', item.get('destinatario', 'Destinatario'))}: ERROR {item.get('error', 'desconocido')}\n"
                )

        thread_ids = resumen.get("thread_ids", [])
        if thread_ids:
            self._append_log("\nThread IDs generados:\n")
            for thread_id in thread_ids:
                self._append_log(f"- {thread_id}\n")

        messagebox.showinfo(
            "Éxito",
            (
                f"Envío completado: {resumen['exitosos']} exitosos, {resumen['fallidos']} fallidos. "
                "Los threadId quedaron registrados en el panel de envío."
            ),
            parent=self.window,
        )

    def _strip_html(self, html_text):
        bloques = []

        for match in re.finditer(r"<(p|ul)\b[^>]*>(.*?)</\1>", html_text, flags=re.IGNORECASE | re.DOTALL):
            etiqueta = match.group(1).lower()
            contenido = match.group(2)

            if etiqueta == "p":
                texto = contenido
                texto = texto.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
                texto = re.sub(r"<[^>]+>", "", texto)
                texto = unescape(texto)
                texto = re.sub(r"[ \t]+", " ", texto)
                texto = re.sub(r"\s*\n\s*", " ", texto)
                texto = re.sub(r"\s{2,}", " ", texto).strip()
                if texto:
                    bloques.append(("p", texto))

            elif etiqueta == "ul":
                items = []
                for li_match in re.finditer(r"<li\b[^>]*>(.*?)</li>", contenido, flags=re.IGNORECASE | re.DOTALL):
                    item = li_match.group(1)
                    item = item.replace("<br>", " ").replace("<br/>", " ").replace("<br />", " ")
                    item = re.sub(r"<[^>]+>", "", item)
                    item = unescape(item)
                    item = re.sub(r"[ \t]+", " ", item)
                    item = re.sub(r"\s*\n\s*", " ", item)
                    item = re.sub(r"\s{2,}", " ", item).strip()
                    if item:
                        items.append(f"• {item}")

                if items:
                    bloques.append(("ul", "\n".join(items)))

        if bloques:
            partes = []
            for indice, (tipo, contenido) in enumerate(bloques):
                if indice > 0:
                    tipo_anterior = bloques[indice - 1][0]
                    if tipo_anterior == "p" and tipo == "ul":
                        separador = "\n"
                    else:
                        separador = "\n\n"
                    partes.append(separador)
                partes.append(contenido)

            texto = "".join(partes)
        else:
            texto = html_text
            texto = texto.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
            texto = re.sub(r"<[^>]+>", "", texto)
            texto = unescape(texto)
            texto = re.sub(r"[ \t]+", " ", texto)
            texto = re.sub(r"\n{3,}", "\n\n", texto)

        return texto.strip()

    def _extraer_resaltados_html(self, html_text):
        if not html_text:
            return []

        resaltados = []
        for match in re.finditer(
            r'<span[^>]*style=["\']([^"\']+)["\'][^>]*>(.*?)</span>',
            html_text,
            flags=re.IGNORECASE | re.DOTALL,
        ):
            estilo = match.group(1).strip()
            contenido = re.sub(r"<[^>]+>", "", match.group(2))
            contenido = unescape(contenido).strip()
            if contenido:
                resaltados.append((contenido, estilo))

        # Evitar duplicados conservando el orden original.
        vistos = set()
        resultado = []
        for texto, estilo in resaltados:
            clave = (texto, estilo)
            if clave in vistos:
                continue
            vistos.add(clave)
            resultado.append((texto, estilo))

        return resultado

    def _aplicar_formato_texto_editable(self, html_base):
        if not html_base:
            return

        # Use the underlying tk.Text for more precise searching and boundary checks
        if hasattr(self.txt_cuerpo, "_textbox"):
            text_widget = self.txt_cuerpo._textbox
        else:
            text_widget = self.txt_cuerpo

        text_widget.tag_remove("resaltado", "1.0", "end")
        fragmentos = [contenido for contenido, _ in self._extraer_resaltados_html(html_base)]

        for fragmento in fragmentos:
            fragmento = fragmento.strip()
            if not fragmento:
                continue

            start_index = "1.0"
            while True:
                found = text_widget.search(fragmento, start_index, stopindex="end")
                if not found:
                    break

                fin = f"{found}+{len(fragmento)}c"

                # Boundary checks: ensure match is not part of a larger token
                char_before = ""
                char_after = ""
                try:
                    if found != "1.0":
                        char_before = text_widget.get(f"{found} -1c", found)
                    char_after = text_widget.get(fin, f"{fin} +1c")
                except Exception:
                    pass

                def is_boundary(ch):
                    return ch == "" or ch.isspace() or ch in ",.;:()\"'–—-"

                if is_boundary(char_before) and is_boundary(char_after):
                    text_widget.tag_add("resaltado", found, fin)
                    break
                else:
                    # continue searching after this match
                    start_index = fin

        # Mantener el cursor al inicio para facilitar la lectura al abrir.
        self.txt_cuerpo.mark_set("insert", "1.0")
        self.txt_cuerpo.see("1.0")

    def _obtener_texto_marcado_desde_editor(self):
        eventos = self.txt_cuerpo._textbox.dump("1.0", "end-1c", text=True, tag=True)
        partes = []

        for tipo, valor, _indice in eventos:
            if tipo == "tagon" and valor == "resaltado":
                partes.append("*")
            elif tipo == "tagoff" and valor == "resaltado":
                partes.append("*")
            elif tipo == "text":
                partes.append(valor)

        return "".join(partes)

    def _aplicar_resaltados_html(self, html_text, resaltados):
        if not html_text or not resaltados:
            return html_text

        # Aplicar primero los textos mas largos para reducir reemplazos parciales.
        ordenados = sorted(resaltados, key=lambda item: len(item[0]), reverse=True)
        salida = html_text

        for texto, estilo in ordenados:
            texto_html = escape(texto)
            reemplazo = f'<span style="{estilo}">{texto_html}</span>'
            salida = re.sub(re.escape(texto_html), reemplazo, salida, count=1)

        return salida

    def _texto_editor_a_html(self, html_base=None):
        """Construye el HTML tomando el texto visible del editor y aplicando
        las etiquetas `resaltado` presentes en el widget. Esto preserva los
        resaltados aunque el usuario edite el texto dentro de ellos."""
        if not hasattr(self.txt_cuerpo, "_textbox"):
            # Fallback: use existing texto->html path via markers
            texto = self.txt_cuerpo.get("1.0", "end").strip()
            return self._texto_a_html(texto, html_base=html_base)

        text_widget = self.txt_cuerpo._textbox
        full_text = text_widget.get("1.0", "end-1c")
        tag_ranges = text_widget.tag_ranges("resaltado")
        if not tag_ranges:
            texto_marcas = full_text
        else:
            partes = []
            cursor = 0
            for i in range(0, len(tag_ranges), 2):
                start_idx = tag_ranges[i]
                end_idx = tag_ranges[i + 1]
                start_off = len(text_widget.get("1.0", start_idx))
                end_off = len(text_widget.get("1.0", end_idx))

                if start_off > cursor:
                    partes.append(full_text[cursor:start_off])
                partes.append(f"*{full_text[start_off:end_off]}*")
                cursor = end_off

            if cursor < len(full_text):
                partes.append(full_text[cursor:])

            texto_marcas = "".join(partes)

        html_resultado = self._texto_a_html(texto_marcas)

        if html_base:
            match_body = re.search(r"<body([^>]*)>", html_base, flags=re.IGNORECASE)
            if match_body:
                body_attrs = match_body.group(1)
                html_resultado = re.sub(
                    r"<body[^>]*>",
                    f"<body{body_attrs}>",
                    html_resultado,
                    count=1,
                    flags=re.IGNORECASE,
                )

        return html_resultado

    def _resaltar_concepto_servicio(self, html_text, estilo_destacado):
        if not html_text:
            return html_text

        patron = r"(El concepto del recibo por honorarios es:\s*)([^<]+)"

        def _reemplazo(match):
            prefijo = match.group(1)
            concepto = match.group(2).strip()
            return f"{prefijo}<span style=\"{estilo_destacado}\">{concepto}</span>"

        return re.sub(patron, _reemplazo, html_text, flags=re.IGNORECASE)

    def _formatear_fragmento_marcado(self, texto, estilo_destacado):
        if not texto:
            return ""

        partes = re.split(r"(\*[^*]+\*)", texto)
        salida = []

        for parte in partes:
            if not parte:
                continue
            if parte.startswith("*") and parte.endswith("*"):
                contenido = escape(parte[1:-1].strip())
                salida.append(f'<span style="{estilo_destacado}">{contenido}</span>')
            else:
                salida.append(escape(parte))

        return "".join(salida)

    def _texto_a_html(self, texto, html_base=None):
        body_attrs = ' style="font-family: Arial; font-size: 11pt;"'
        resaltados = []
        if html_base:
            match_body = re.search(r"<body([^>]*)>", html_base, flags=re.IGNORECASE)
            if match_body:
                body_attrs = match_body.group(1)
            resaltados = self._extraer_resaltados_html(html_base)

        estilo_destacado = "font-weight: bold; color: #073763;"
        for _, estilo in resaltados:
            if "color" in estilo.lower() or "font-weight" in estilo.lower():
                estilo_destacado = estilo
                break


        lineas = texto.splitlines()
        bloques_html = []
        lista_actual = []
        parrafo_actual = []

        def flush_parrafo():
            nonlocal parrafo_actual
            if not parrafo_actual:
                return
            contenido = "<br>".join(self._formatear_fragmento_marcado(l, estilo_destacado) for l in parrafo_actual)
            bloques_html.append(f"<p>{contenido}</p>")
            parrafo_actual = []

        def flush_lista():
            nonlocal lista_actual
            if not lista_actual:
                return
            items = "".join(f'<li style="margin:0;padding:0">{self._formatear_fragmento_marcado(item, estilo_destacado)}</li>' for item in lista_actual)
            # Inline styles to force indentation in mail clients
            bloques_html.append(f'<ul style="margin:0 0 6px 22px; padding-left:18px;">{items}</ul>')
            lista_actual = []

        for raw in lineas:
            linea = raw.rstrip()
            if not linea.strip():
                # Blank line separates paragraphs/lists
                flush_parrafo()
                flush_lista()
                continue

            m = re.match(r"^(?:\s*)(?:[-•])\s*(.+)$", linea)
            if m:
                # It's a list item
                item_text = m.group(1).strip()
                # if we had paragraph content, flush it first
                if parrafo_actual:
                    flush_parrafo()
                lista_actual.append(item_text)
            else:
                # Normal paragraph line
                # if we have an active list, flush it before adding paragraph lines
                if lista_actual:
                    flush_lista()
                parrafo_actual.append(linea.strip())

        # flush remaining
        flush_parrafo()
        flush_lista()

        if not bloques_html:
            bloques_html = ["<p></p>"]

        html_resultado = f"<html><body{body_attrs}>" + "".join(bloques_html)
        html_resultado += "</body></html>"

        html_resultado = self._aplicar_resaltados_html(html_resultado, resaltados)
        return self._resaltar_concepto_servicio(html_resultado, estilo_destacado)

    def _guardar_estado_actual(self):
        if not self.data or not self._form_loaded:
            return
        item = self.data[self.idx_correo]
        item["asunto"] = self.entry_asunto.get().strip()
        texto_visible = self.txt_cuerpo.get("1.0", "end").strip()
        item["cuerpo_texto"] = texto_visible
        texto_marcado = self._obtener_texto_marcado_desde_editor().strip()
        if texto_visible == item.get("cuerpo_texto_original", ""):
            item["cuerpo_html"] = item.get("cuerpo_html_original", item.get("cuerpo_html", ""))
        else:
            html_base = item.get("cuerpo_html_original", item.get("cuerpo_html", ""))

            # Build HTML directly from editor while preserving current tag ranges
            item["cuerpo_html"] = self._texto_editor_a_html(html_base=html_base)
            # Update originals so subsequent loads use the edited version
            item["cuerpo_html_original"] = item["cuerpo_html"]
            item["cuerpo_texto_original"] = texto_visible

    def _obtener_pdf_actual(self):
        if not self.pdf_paths_actuales:
            return ""
        idx = max(0, min(self.idx_adjunto, len(self.pdf_paths_actuales) - 1))
        return self.pdf_paths_actuales[idx]

    def _actualizar_navegacion_adjuntos(self):
        total = len(self.pdf_paths_actuales)
        if total == 0:
            self.lbl_adjunto.configure(text="Adjunto 0 de 0")
            self.btn_prev_adjunto.configure(state="disabled")
            self.btn_next_adjunto.configure(state="disabled")
            return

        self.lbl_adjunto.configure(text=f"Adjunto {self.idx_adjunto + 1} de {total}")
        self.btn_prev_adjunto.configure(state="normal" if self.idx_adjunto > 0 else "disabled")
        self.btn_next_adjunto.configure(state="normal" if self.idx_adjunto < total - 1 else "disabled")

    def _mostrar_correo(self, idx):
        if not self.data:
            return
        self._guardar_estado_actual()

        idx = max(0, min(idx, len(self.data) - 1))
        self.idx_correo = idx
        self.idx_adjunto = 0
        self.idx_pagina_pdf = 0
        self.zoom_factor = 2.4
        self.zoom_slider.set(self.zoom_factor)
        self._actualizar_zoom_label()

        item = self.data[idx]
        self.pdf_paths_actuales = item.get("pdf_paths") or [item.get("pdf_path", "")]
        self._actualizar_navegacion_adjuntos()

        self.lbl_correo.configure(text=f"Correo {idx + 1} de {len(self.data)}")
        self.lbl_dest.configure(text=f"Para: {item.get('nombre', '')} <{item.get('destinatario', '')}>")
        self.btn_prev_correo.configure(state="normal" if idx > 0 else "disabled")
        self.btn_next_correo.configure(state="normal" if idx < len(self.data) - 1 else "disabled")

        self.entry_asunto.delete(0, "end")
        self.entry_asunto.insert(0, item.get("asunto", ""))

        self.txt_cuerpo.delete("1.0", "end")
        self.txt_cuerpo.insert("1.0", item.get("cuerpo_texto", ""))
        self._aplicar_formato_texto_editable(item.get("cuerpo_html_original", item.get("cuerpo_html", "")))
        self._form_loaded = True

        self._renderizar_pdf(self._obtener_pdf_actual(), 0)

    def _renderizar_pdf(self, pdf_path, pagina=0):
        try:
            if not pdf_path or not os.path.exists(pdf_path):
                raise FileNotFoundError("No se encontró el PDF adjunto")

            pdf = pdfium.PdfDocument(pdf_path)
            self.total_paginas_pdf = len(pdf)
            pagina = max(0, min(pagina, self.total_paginas_pdf - 1))
            self.idx_pagina_pdf = pagina
            page = pdf[pagina]

            self.canvas_pdf.update_idletasks()
            cw = max(self.canvas_pdf.winfo_width(), 320)
            ch = max(self.canvas_pdf.winfo_height(), 420)

            pw, ph = page.get_size()
            fit_scale = min(cw / pw, ch / ph)
            scale = max(0.2, min(fit_scale * self.zoom_factor, 5.0))

            bitmap = page.render(scale=scale)
            pil_img = bitmap.to_pil()
            self._img_ref = ImageTk.PhotoImage(pil_img)

            self.canvas_pdf.delete("all")
            self._pdf_image_id = self.canvas_pdf.create_image(0, 0, image=self._img_ref, anchor="nw")
            self.canvas_pdf.configure(scrollregion=(0, 0, pil_img.width, pil_img.height))
            self.canvas_pdf.xview_moveto(0.0)
            self.canvas_pdf.yview_moveto(0.0)

            self.lbl_pagina.configure(text=f"Página {pagina + 1} de {self.total_paginas_pdf}")
            self.btn_prev_pagina.configure(state="normal" if pagina > 0 else "disabled")
            self.btn_next_pagina.configure(state="normal" if pagina < self.total_paginas_pdf - 1 else "disabled")
            pdf.close()
        except Exception as e:
            self.canvas_pdf.delete("all")
            cw = max(self.canvas_pdf.winfo_width(), 320)
            ch = max(self.canvas_pdf.winfo_height(), 420)
            self.canvas_pdf.create_text(cw // 2, ch // 2, text=f"Error al cargar PDF:\n{e}", width=cw - 30, fill="#B00020")
            self.canvas_pdf.configure(scrollregion=(0, 0, cw, ch))
            self.lbl_pagina.configure(text="Página 0 de 0")
            self.btn_prev_pagina.configure(state="disabled")
            self.btn_next_pagina.configure(state="disabled")

    def _actualizar_zoom_label(self):
        self.lbl_zoom.configure(text=f"{int(self.zoom_factor * 100)}%")

    def _cambiar_zoom(self, delta):
        nuevo_zoom = max(0.5, min(self.zoom_factor + delta, 3.0))
        self.zoom_factor = nuevo_zoom
        self.zoom_slider.set(nuevo_zoom)
        self._actualizar_zoom_label()
        if self.data:
            self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf)

    def _on_zoom_slider(self, value):
        self.zoom_factor = float(value)
        self._actualizar_zoom_label()
        if self.data:
            self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf)

    def _reset_zoom(self):
        self.zoom_factor = 2.4
        self.zoom_slider.set(2.4)
        self._actualizar_zoom_label()
        if self.data:
            self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf)

    def _ctrl_presionado(self, event):
        # Tk usa bitmask en event.state; 0x0004 representa Control en Windows/X11.
        return bool(event.state & 0x0004)

    def _on_mouse_wheel(self, event):
        if not self._ctrl_presionado(event):
            # Scroll libre del PDF cuando no se presiona Ctrl.
            if getattr(event, "num", None) == 4:
                self.canvas_pdf.yview_scroll(-1, "units")
                return "break"
            if getattr(event, "num", None) == 5:
                self.canvas_pdf.yview_scroll(1, "units")
                return "break"

            if hasattr(event, "delta") and event.delta:
                pasos = -1 if event.delta > 0 else 1
                if event.state & 0x0001:
                    self.canvas_pdf.xview_scroll(pasos * 3, "units")
                else:
                    self.canvas_pdf.yview_scroll(pasos * 3, "units")
                return "break"
            return

        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = 1 if event.delta > 0 else -1
        elif getattr(event, "num", None) == 4:
            delta = 1
        elif getattr(event, "num", None) == 5:
            delta = -1

        if delta == 0:
            return "break"

        # Paso suave para evitar cambios bruscos por notch.
        self._cambiar_zoom(0.08 if delta > 0 else -0.08)
        return "break"

    def _on_pan_start(self, event):
        self.canvas_pdf.scan_mark(event.x, event.y)

    def _on_pan_drag(self, event):
        # Pan continuo arrastrando con click izquierdo.
        self.canvas_pdf.scan_dragto(event.x, event.y, gain=1)

    def _anterior_correo(self):
        if self.idx_correo > 0:
            self._mostrar_correo(self.idx_correo - 1)

    def _siguiente_correo(self):
        if self.idx_correo < len(self.data) - 1:
            self._mostrar_correo(self.idx_correo + 1)

    def _anterior_pagina(self):
        self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf - 1)

    def _siguiente_pagina(self):
        self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf + 1)

    def _anterior_adjunto(self):
        if self.idx_adjunto > 0:
            self.idx_adjunto -= 1
            self.idx_pagina_pdf = 0
            self._actualizar_navegacion_adjuntos()
            self._renderizar_pdf(self._obtener_pdf_actual(), 0)

    def _siguiente_adjunto(self):
        if self.idx_adjunto < len(self.pdf_paths_actuales) - 1:
            self.idx_adjunto += 1
            self.idx_pagina_pdf = 0
            self._actualizar_navegacion_adjuntos()
            self._renderizar_pdf(self._obtener_pdf_actual(), 0)

    def _on_pdf_resize(self, _event):
        if not self.data:
            return
        if self._render_job is not None:
            self.window.after_cancel(self._render_job)
        self._render_job = self.window.after(
            120,
            lambda: self._renderizar_pdf(self._obtener_pdf_actual(), self.idx_pagina_pdf),
        )

    def _guardar(self):
        self._guardar_estado_actual()
        if self.on_save:
            self.on_save(copy.deepcopy(self.data))
        messagebox.showinfo("Éxito", "Cambios guardados en la previsualización", parent=self.window)

    def _enviar(self):
        self._guardar_estado_actual()
        if not self.on_send:
            return
        try:
            self._limpiar_logs()
            self.lbl_estado_envio.configure(text="Enviando...", text_color="#FFD700")
            self.btn_enviar.configure(state="disabled")
            self.on_send(copy.deepcopy(self.data), self._append_log, self._finalizar_envio)
        except Exception as e:
            self.btn_enviar.configure(state="normal")
            messagebox.showerror("Error", str(e), parent=self.window)

    def _cerrar(self):
        self.window.grab_release()
        self.window.destroy()
