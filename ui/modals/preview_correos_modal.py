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
        self.idx_pagina_pdf = 0
        self.total_paginas_pdf = 1
        self._img_ref = None
        self._render_job = None
        self._form_loaded = False
        self.zoom_factor = 1.0
        self._platform = platform.system()
        self._pdf_image_id = None

        for item in self.data:
            cuerpo_texto = self._strip_html(item.get("cuerpo_html", ""))
            item.setdefault("cuerpo_texto", cuerpo_texto)
            item.setdefault("cuerpo_texto_original", cuerpo_texto)
            item.setdefault("cuerpo_html_original", item.get("cuerpo_html", ""))

        self.window = ctk.CTkToplevel(parent)
        self.window.title("Previsualizar correos")
        self.window.geometry("1300x760")
        self.window.state("zoomed")
        self.window.minsize(1100, 650)
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
        split.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)

        panel_pdf = ctk.CTkFrame(split, fg_color=CARD_COLOR, corner_radius=10)
        panel_pdf.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        ctk.CTkLabel(panel_pdf, text="PDF adjunto", text_color=TEXT_COLOR, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

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
        self.zoom_slider.set(1.0)
        self.zoom_slider.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.lbl_zoom = ctk.CTkLabel(zoom_frame, text="100%", text_color=TEXT_LIGHT, width=56)
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

        footer = ctk.CTkFrame(self.window, fg_color=CARD_COLOR)
        footer.pack(fill="x", padx=12, pady=(0, 12))

        btns = ctk.CTkFrame(footer, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(btns, text="Cerrar", fg_color=BG_COLOR, text_color=TEXT_COLOR, command=self._cerrar).pack(side="left")
        ctk.CTkButton(btns, text="Guardar cambios", fg_color=PRIMARY_COLOR, hover_color=ACCENT_COLOR, command=self._guardar).pack(side="right", padx=(8, 0))
        ctk.CTkButton(btns, text="Enviar correos", fg_color=ACCENT_COLOR, hover_color=PRIMARY_COLOR, command=self._enviar).pack(side="right")

    def _strip_html(self, html_text):
        texto = html_text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
        texto = texto.replace("</p>", "\n\n").replace("</li>", "\n")
        texto = re.sub(r"<li[^>]*>", "- ", texto, flags=re.IGNORECASE)
        texto = re.sub(r"<[^>]+>", "", texto)
        texto = unescape(texto)
        texto = re.sub(r"\n{3,}", "\n\n", texto)
        return texto.strip()

    def _texto_a_html(self, texto):
        partes = [f"<p>{escape(linea)}</p>" for linea in texto.split("\n\n") if linea.strip()]
        if not partes:
            return "<html><body><p></p></body></html>"
        return "<html><body style=\"font-family: Arial; font-size: 11pt;\">" + "".join(partes) + "</body></html>"

    def _guardar_estado_actual(self):
        if not self.data or not self._form_loaded:
            return
        item = self.data[self.idx_correo]
        item["asunto"] = self.entry_asunto.get().strip()
        texto = self.txt_cuerpo.get("1.0", "end").strip()
        item["cuerpo_texto"] = texto
        if texto == item.get("cuerpo_texto_original", ""):
            item["cuerpo_html"] = item.get("cuerpo_html_original", item.get("cuerpo_html", ""))
        else:
            item["cuerpo_html"] = self._texto_a_html(texto)

    def _mostrar_correo(self, idx):
        if not self.data:
            return
        self._guardar_estado_actual()

        idx = max(0, min(idx, len(self.data) - 1))
        self.idx_correo = idx
        self.idx_pagina_pdf = 0
        self.zoom_factor = 1.0
        self.zoom_slider.set(self.zoom_factor)
        self._actualizar_zoom_label()

        item = self.data[idx]

        self.lbl_correo.configure(text=f"Correo {idx + 1} de {len(self.data)}")
        self.lbl_dest.configure(text=f"Para: {item.get('nombre', '')} <{item.get('destinatario', '')}>")
        self.btn_prev_correo.configure(state="normal" if idx > 0 else "disabled")
        self.btn_next_correo.configure(state="normal" if idx < len(self.data) - 1 else "disabled")

        self.entry_asunto.delete(0, "end")
        self.entry_asunto.insert(0, item.get("asunto", ""))

        self.txt_cuerpo.delete("1.0", "end")
        self.txt_cuerpo.insert("1.0", item.get("cuerpo_texto", ""))
        self._form_loaded = True

        self._renderizar_pdf(item.get("pdf_path", ""), 0)

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
            self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf)

    def _on_zoom_slider(self, value):
        self.zoom_factor = float(value)
        self._actualizar_zoom_label()
        if self.data:
            self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf)

    def _reset_zoom(self):
        self.zoom_factor = 1.0
        self.zoom_slider.set(1.0)
        self._actualizar_zoom_label()
        if self.data:
            self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf)

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
        self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf - 1)

    def _siguiente_pagina(self):
        self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf + 1)

    def _on_pdf_resize(self, _event):
        if not self.data:
            return
        if self._render_job is not None:
            self.window.after_cancel(self._render_job)
        self._render_job = self.window.after(
            120,
            lambda: self._renderizar_pdf(self.data[self.idx_correo].get("pdf_path", ""), self.idx_pagina_pdf),
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
            self.on_send(copy.deepcopy(self.data))
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self.window)

    def _cerrar(self):
        self.window.grab_release()
        self.window.destroy()
